import easyocr
import cv2
import pandas as pd
import numpy as np
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox

class BeamDigitizer:
    def __init__(self):
        self.reader = None

    def get_reader(self):
        if self.reader is None:
            self.reader = easyocr.Reader(['en'])
        return self.reader

    def extract_table_data(self, image_path):
        img = cv2.imread(image_path)
        if img is None: raise FileNotFoundError(f"Could not read image: {image_path}")
        results = self.get_reader().readtext(image_path)
        table_data = []
        for (bbox, text, prob) in results:
            x_min = min(p[0] for p in bbox)
            if x_min < 350: table_data.append({'text': text, 'bbox': bbox, 'prob': prob})
        return table_data

    def parse_parameters(self, table_data):
        rows = {}
        for item in table_data:
            y_center = sum(p[1] for p in item['bbox']) / 4
            row_key = round(y_center / 12) * 12
            if row_key not in rows: rows[row_key] = []
            rows[row_key].append(item)
        parsed_params = []
        for r_key in sorted(rows.keys()):
            row_items = sorted(rows[r_key], key=lambda x: x['bbox'][0][0])
            line_text = " ".join([it['text'] for it in row_items])
            line_text = line_text.replace(',', '.')
            line_text = line_text.replace('Is0', 'ISO').replace('IS0', 'ISO')

            # Refined regex to better capture labels and values
            match = re.search(r'([A-Za-z\^2\s\*\"]+)\s+([\d\.]+)', line_text)
            if match:
                label = match.group(1).strip()
                value = match.group(2).strip()
                # Skip header-like rows that were misidentified as label-value
                if len(label) > 1:
                    parsed_params.append({'Parameter': label, 'Value': value})
            else:
                # Store lines that might be headers or groups
                if len(line_text.strip()) > 3:
                    parsed_params.append({'Parameter': line_text.strip(), 'Value': ''})

        return parsed_params

    def extract_graph_data(self, image_path):
        img = cv2.imread(image_path)
        graph_area = img[50:550, 650:]
        hsv = cv2.cvtColor(graph_area, cv2.COLOR_BGR2HSV)

        lower_orange = np.array([0, 100, 100])
        upper_orange = np.array([25, 255, 255])
        mask_x = cv2.inRange(hsv, lower_orange, upper_orange)

        lower_blue = np.array([100, 100, 100])
        upper_blue = np.array([130, 255, 255])
        mask_y = cv2.inRange(hsv, lower_blue, upper_blue)

        def find_points(mask, label):
            contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            points = []
            for cnt in contours:
                M = cv2.moments(cnt)
                if M["m00"] != 0:
                    cx = int(M["m10"] / M["m00"])
                    cy = int(M["m01"] / M["m00"])
                    z_val = 400 + (cx + 650 - 780) * (600 - 400) / (1020 - 780)
                    w_val = 180 + (cy + 50 - 460) * (390 - 180) / (110 - 460)
                    points.append({'Type': label, 'Z Location (mm)': round(z_val, 2), 'Beam Width (um)': round(w_val, 2)})
            return points

        points_x = find_points(mask_x, 'X')
        points_y = find_points(mask_y, 'Y')
        # Sort by Z Location for readability in Excel
        all_points = sorted(points_x + points_y, key=lambda x: (x['Type'], x['Z Location (mm)']))
        return all_points

    def save_to_excel(self, params, points, output_path):
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_params = pd.DataFrame(params)
            df_params.to_excel(writer, sheet_name='Beam Parameters', index=False)
            df_points = pd.DataFrame(points)
            df_points.to_excel(writer, sheet_name='Graph Data', index=False)
        return output_path

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("ビームプロファイラ画像数値化ツール")
        self.digitizer = BeamDigitizer()

        self.label = tk.Label(root, text="解析する画像を選択してください")
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="画像を選択", command=self.select_image)
        self.select_button.pack(pady=5)

        self.file_path = ""
        self.path_label = tk.Label(root, text="")
        self.path_label.pack(pady=5)

        self.run_button = tk.Button(root, text="解析実行 & Excel保存", command=self.run_analysis, state=tk.DISABLED)
        self.run_button.pack(pady=20)

    def select_image(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp")])
        if self.file_path:
            self.path_label.config(text=os.path.basename(self.file_path))
            self.run_button.config(state=tk.NORMAL)

    def run_analysis(self):
        try:
            self.label.config(text="解析中... しばらくお待ちください")
            self.root.update()

            params = self.digitizer.parse_parameters(self.digitizer.extract_table_data(self.file_path))
            points = self.digitizer.extract_graph_data(self.file_path)

            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if output_path:
                self.digitizer.save_to_excel(params, points, output_path)
                messagebox.showinfo("完了", f"解析結果が保存されました:\n{output_path}")

            self.label.config(text="解析する画像を選択してください")
        except Exception as e:
            messagebox.showerror("エラー", f"解析中にエラーが発生しました:\n{str(e)}")
            self.label.config(text="エラーが発生しました")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        d = BeamDigitizer()
        p = d.parse_parameters(d.extract_table_data(sys.argv[1]))
        pts = d.extract_graph_data(sys.argv[1])
        out = "output.xlsx"
        d.save_to_excel(p, pts, out)
        print(f"CLI Mode: Saved to {out}")
    else:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
