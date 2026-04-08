from issue_aggregator import extract_data, append_data
import os

def test():
    source_files = ["test_individual_1.xlsx", "test_individual_2.xlsx"]
    summary_file = "test_summary.xlsx"

    all_data = []
    for f in source_files:
        print(f"Extracting from {f}...")
        data = extract_data(f)
        if data:
            all_data.append(data)
            print(f"  Data: {data}")

    if all_data:
        print(f"Appending {len(all_data)} items to {summary_file}...")
        success = append_data(summary_file, all_data)
        if success:
            print("Aggregation successful!")
        else:
            print("Aggregation failed!")
    else:
        print("No data extracted.")

if __name__ == "__main__":
    test()
