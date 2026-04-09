import easyocr
import cv2
import pandas as pd
import numpy as np
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from datetime import datetime

class LaserQualityManager:
    def __init__(self):
        self.reader = None

    def get_reader(self):
        if self.reader is None:
            # First time might take a while to download/load
            self.reader = easyocr.Reader(['en'])
        return self.reader

    def extract_data_from_image(self, image_path):
        """Extracts parameters and graph points from a single image."""
        reader = self.get_reader()
        results = reader.readtext(image_path)

        # Filter for table data (usually left side)
        table_data = []
        for (bbox, text, prob) in results:
            x_min = min(p[0] for p in bbox)
            if x_min < 400: # Slightly wider than original beam_digitizer
                table_data.append({'text': text, 'bbox': bbox, 'prob': prob})

        params = self.parse_parameters(table_data)
        return params

    def parse_parameters(self, table_data):
        rows = {}
        for item in table_data:
            y_center = sum(p[1] for p in item['bbox']) / 4
            row_key = round(y_center / 12) * 12
            if row_key not in rows: rows[row_key] = []
            rows[row_key].append(item)

        parsed_params = {}
        for r_key in sorted(rows.keys()):
            row_items = sorted(rows[r_key], key=lambda x: x['bbox'][0][0])
            line_text = " ".join([it['text'] for it in row_items])
            line_text = line_text.replace(',', '.')
            line_text = line_text.replace('Is0', 'ISO').replace('IS0', 'ISO')
            line_text = line_text.replace('M ^ 2', 'M^2').replace('M ^2', 'M^2').replace('M^ 2', 'M^2')

            # Look for Label followed by Number
            match = re.search(r'([A-Za-z\^2\d\.\s\*\"]+?)\s+([\d\.]+)', line_text)
            if match:
                label = match.group(1).strip()
                value = match.group(2).strip()
                if len(label) > 1:
                    try:
                        parsed_params[label] = float(value)
                    except ValueError:
                        parsed_params[label] = value
        return parsed_params

    def calculate_cpk(self, data, usl, lsl):
        if len(data) < 2: return None, None
        mean = np.mean(data)
        std = np.std(data, ddof=1)
        if std == 0: return None, None

        cp = (usl - lsl) / (6 * std)
        cpu = (usl - mean) / (3 * std)
        cpl = (mean - lsl) / (3 * std)
        cpk = min(cpu, cpl)
        return round(cp, 2), round(cpk, 2)

    def process_batch(self, image_paths, master_excel_path):
        all_results = []
        for path in image_paths:
            params = self.extract_data_from_image(path)
            res = {'Filename': os.path.basename(path), 'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            res.update(params)
            all_results.append(res)

        if not all_results:
            return None

        df = pd.DataFrame(all_results)

        # Analysis Sheet
        analysis_data = []
        # Key parameters to analyze.
        # We try to find common profiler labels
        target_params = ['M^2', 'M2', 'Waist Width', 'Waist Location', 'Divergence']

        # Example Spec Limits (USL, LSL) - In production, these should be configurable
        # Here we provide some dummy specs for demonstration of the Cp/Cpk calculation
        specs = {
            'M^2': (1.5, 1.0),
            'M2': (1.5, 1.0),
            'Waist Width': (300, 100),
            'Divergence': (15, 0)
        }

        for p in df.columns:
            # Check if this column is one of our targets or contains a target string
            is_target = any(t.lower() in p.lower() for t in target_params)
            if is_target:
                numeric_data = pd.to_numeric(df[p], errors='coerce').dropna()
                if not numeric_data.empty and len(numeric_data) >= 2:
                    # Find matching spec
                    spec_key = next((t for t in specs if t.lower() in p.lower()), None)
                    usl, lsl = specs.get(spec_key, (None, None))

                    cp, cpk = (None, None)
                    if usl is not None and lsl is not None:
                        cp, cpk = self.calculate_cpk(numeric_data, usl, lsl)

                    analysis_data.append({
                        'Parameter': p,
                        'Count': len(numeric_data),
                        'Mean': round(float(numeric_data.mean()), 3),
                        'StdDev': round(float(numeric_data.std()), 3),
                        'Max': numeric_data.max(),
                        'Min': numeric_data.min(),
                        'USL': usl,
                        'LSL': lsl,
                        'Cp': cp,
                        'Cpk': cpk
                    })

        df_analysis = pd.DataFrame(analysis_data)

        with pd.ExcelWriter(master_excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Raw Data', index=False)
            if not df_analysis.empty:
                df_analysis.to_excel(writer, sheet_name='Process Capability', index=False)

        return master_excel_path

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("レーザ品質管理システム (Laser Quality Manager)")
        self.root.geometry("600x450")
        self.manager = LaserQualityManager()

        self.image_files = []

        tk.Label(root, text="レーザ発振器 品質管理・分析ツール", font=("", 14, "bold")).pack(pady=10)

        tk.Label(root, text="1. 解析する画像フォルダまたはファイルを選択してください", font=("", 10)).pack(pady=5)

        btn_frame = tk.Frame(root)
        btn_frame.pack()
        tk.Button(btn_frame, text="ファイルを選択", command=self.select_files).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="フォルダを選択", command=self.select_folder).pack(side=tk.LEFT, padx=5)

        self.lbl_status = tk.Label(root, text="選択されていません", fg="gray", wraplength=500)
        self.lbl_status.pack(pady=5)

        tk.Label(root, text="2. 集計・分析を実行", font=("", 10)).pack(pady=5)
        self.run_button = tk.Button(root, text="一括解析 & Excel保存", command=self.run_analysis,
                                    bg="lightblue", state=tk.DISABLED, height=2, width=25)
        self.run_button.pack(pady=10)

        self.progress_label = tk.Label(root, text="")
        self.progress_label.pack()

        # Instruction label
        instructions = (
            "※ M^2, Waist Width, Divergence 等の主要パラメータを自動抽出します。\n"
            "※ 集計後のExcelには、平均・標準偏差・工程管理能力(Cp/Cpk)が計算されます。"
        )
        tk.Label(root, text=instructions, justify=tk.LEFT, fg="#555").pack(pady=20)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp")])
        if files:
            self.image_files = list(files)
            self.lbl_status.config(text=f"{len(self.image_files)} 個のファイルを選択中", fg="black")
            self.run_button.config(state=tk.NORMAL)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            files = [os.path.join(folder, f) for f in os.listdir(folder)
                     if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))]
            if files:
                self.image_files = files
                self.lbl_status.config(text=f"フォルダ内 {len(self.image_files)} 個の画像を選択中", fg="black")
                self.run_button.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("警告", "指定されたフォルダに画像が見つかりませんでした。")

    def run_analysis(self):
        if not self.image_files:
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel Files", "*.xlsx")],
                                                 initialfile=f"Laser_Quality_Report_{datetime.now().strftime('%Y%m%d')}.xlsx")
        if not output_path:
            return

        try:
            self.run_button.config(state=tk.DISABLED)
            self.progress_label.config(text="解析中... OCR処理には時間がかかります")
            self.root.update()

            self.manager.process_batch(self.image_files, output_path)

            messagebox.showinfo("完了", f"解析と集計が完了しました。\n保存先: {output_path}")
            self.progress_label.config(text="完了")
        except Exception as e:
            messagebox.showerror("エラー", f"解析中にエラーが発生しました:\n{str(e)}")
            self.progress_label.config(text="エラー発生")
        finally:
            self.run_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    # Support for CLI mode for verification
    import sys
    if len(sys.argv) > 1:
        img_path = sys.argv[1]
        out_path = "batch_output.xlsx"
        if os.path.isdir(img_path):
            imgs = [os.path.join(img_path, f) for f in os.listdir(img_path) if f.lower().endswith(('.png', '.jpg'))]
        else:
            imgs = [img_path]

        m = LaserQualityManager()
        m.process_batch(imgs, out_path)
        print(f"Batch processing complete. Results saved to {out_path}")
    else:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
