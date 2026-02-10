import pandas as pd
import pathlib
import re
import argparse
import subprocess
import sys
import io

def get_excel_file_paths(root_path):
    """
    Traverse the directory tree and return absolute paths of matching files.
    """
    root = pathlib.Path(root_path)
    if not root.exists() or not root.is_dir():
        print(f"Error: Directory '{root_path}' not found.")
        return []

    l1_pattern = re.compile(r"^(\d{2})-(\d{2})-(\d{2})$")
    file_list = []

    for path_l1 in sorted(root.iterdir()):
        match_l1 = l1_pattern.match(path_l1.name)
        
        if path_l1.is_dir() and match_l1:
            dd, mm, yy = match_l1.groups()
            expected_date = f"{dd}{mm}20{yy}"
            path_l2 = path_l1 / f"{dd}.{mm}"
            
            if path_l2.exists() and path_l2.is_dir():
                for file_path in path_l2.iterdir():
                    if file_path.is_file():
                        if file_path.name.startswith(f"PA_Nhapkhau24H_{expected_date}"):
                            file_list.append(str(file_path.absolute()))
    return file_list

def main():
    # 1. Thiết lập đối số dòng lệnh
    parser = argparse.ArgumentParser(description="Scan folder and collect data using subprocess.")
    parser.add_argument("folder_path", help="Path to the root folder to scan")
    args = parser.parse_args()

    # 2. Lấy danh sách file excel
    print(f"Scanning folder: {args.folder_path}...")
    excel_files = get_excel_file_paths(args.folder_path)

    if not excel_files:
        print("No matching Excel files found.")
        return

    print(f"Found {len(excel_files)} files. Starting data extraction...")

    all_dataframes = []

    # 3. Duyệt qua từng file và gọi subprocess
    for file_path in excel_files:
        try:
            print(f"Processing: {pathlib.Path(file_path).name}")
            
            # Gọi tiến trình CollectDataPlot.py
            # Lưu ý: CollectDataPlot.py cần được sửa để in kết quả ra CSV (xem lưu ý bên dưới)
            result = subprocess.run(
                [sys.executable, "CollectDataPlot.py", file_path],
                capture_output=True,
                text=True,
                encoding='utf-8',
                check=True
            )

            # Đọc dữ liệu từ stdout của tiến trình con
            if result.stdout.strip():
                # Giả định script con in ra dữ liệu dạng CSV
                df_temp = pd.read_csv(io.StringIO(result.stdout))
                all_dataframes.append(df_temp)
            
        except subprocess.CalledProcessError as e:
            print(f"Error processing {file_path}: {e.stderr}")
        except Exception as e:
            print(f"An error occurred: {e}")

    # 4. Ghép các dataframe vào dataframe tổng
    if all_dataframes:
        final_df = pd.concat(all_dataframes, ignore_index=True)
        
        # Sắp xếp lại TT (Thứ tự) cho toàn bộ bảng tổng
        final_df['TT'] = range(1, len(final_df) + 1)
        
        print("\n--- Final Aggregated Data ---")
        print(final_df.to_string(index=False))
        
        # Lưu kết quả tổng hợp
        final_df.to_csv("Aggregated_Power_Data.csv", index=False)
        print(f"\nSaved aggregated data to 'Aggregated_Power_Data.csv'")
    else:
        print("No data collected.")

if __name__ == "__main__":
    main()