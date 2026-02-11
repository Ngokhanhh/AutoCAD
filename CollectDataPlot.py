import pandas as pd
import argparse
import os
import sys
import re

def main():
    # 1. Cấu hình đối số dòng lệnh
    parser = argparse.ArgumentParser(description="Program to extract and reshape power data from Excel.")
    parser.add_argument("input_file", help="Path to the input Excel file")
    
    args = parser.parse_args()
    file_path = args.input_file

    # Kiểm tra file tồn tại
    if not os.path.exists(file_path):
        print(f"Error: File not found at path: {file_path}")
        sys.exit(1)

    # --- Bước trích xuất ngày từ tên file ---
    # Format: PA_Nhapkhau24H_ddmmyyyy_... -> Lấy cụm 8 chữ số
    file_name = os.path.basename(file_path)
    date_match = re.search(r'(\d{2})(\d{2})(\d{4})', file_name)
    if date_match:
        day, month, year = date_match.groups()
        date_str = f"{day}/{month}/{year}"
    else:
        date_str = "01/01/2026"  # Giá trị mặc định nếu không tìm thấy
        print("Warning: Could not extract date from filename, using default.")

    try:
        # 2. Đọc file raw để xác định hàng tiêu đề
        df_raw = pd.read_excel(file_path, header=None)

        header_row_index = None
        for i, row in df_raw.iterrows():
            if "TBA" in row.values and "0h:30" in row.values:
                header_row_index = i
                break

        if header_row_index is None:
            print("Error: Could not find header row with 'TBA' and '0h:30'!")
            return

        # Đọc lại file với header chuẩn
        df = pd.read_excel(file_path, header=header_row_index)

        # 3. Định nghĩa các nhãn cần lấy
        target_labels = ['E20.3', 'E22.4', 'E5.7']
        start_col = '0h:30'
        end_col = '24h'

        # Xác định phạm vi cột thời gian
        all_cols = df.columns.tolist()
        try:
            start_idx = all_cols.index(start_col)
            # Kiểm tra nếu end_col tồn tại, nếu không lấy đến cột cuối cùng có định dạng thời gian
            end_idx = all_cols.index(end_col) if end_col in all_cols else len(all_cols) - 1
            time_columns = all_cols[start_idx : end_idx + 1]
        except ValueError:
            print(f"Error: Could not find time columns from {start_col} to {end_col}")
            return

        # 4. Lọc dữ liệu theo target_labels và tính tổng (nếu một TBA có nhiều đường dây)
        df_filtered = df[df['TBA'].isin(target_labels)].copy()
        
        # Nhóm theo TBA và cộng các giá trị công suất
        df_grouped = df_filtered.groupby('TBA')[time_columns].sum()

        # 5. Chuyển đổi bảng (Transpose) để đưa thời gian về dạng hàng
        df_reshaped = df_grouped.transpose()
        df_reshaped.reset_index(inplace=True)
        df_reshaped.rename(columns={'index': 'Time_Raw'}, inplace=True)

        # 6. Tạo cột Time chuẩn (hh:mm dd/mm/yyyy)
        def format_time(row_time):
            # Loại bỏ chữ 'h', khoảng trắng và đưa về dạng chuỗi
            t_str = str(row_time).lower().replace('h', '').strip()
            
            if ':' in t_str:
                # Trường hợp 0h:30, 13:30
                h, m = t_str.split(':')
            else:
                # Trường hợp giờ tròn như 1h, 2, 13
                h, m = t_str, "00"
            
            # Ép kiểu int để format 2 chữ số (ví dụ: 1 -> 01)
            try:
                return f"{int(h):02d}:{int(m):02d} {date_str}"
            except:
                return f"{t_str} {date_str}" # Tránh crash nếu dữ liệu lạ

        df_reshaped['Time'] = df_reshaped['Time_Raw'].apply(format_time)

        # 7. Sắp xếp lại thứ tự cột và thêm cột TT (Thứ tự)
        final_cols = ['Time'] + [label for label in target_labels if label in df_reshaped.columns]
        df_final = df_reshaped[final_cols].copy()
        df_final.insert(0, 'TT', range(1, len(df_final) + 1))

        # Hiển thị kết quả
        # print("\n--- Data Extracted Successfully ---")
        # print(df_final.to_string(index=False))
        df_final.to_csv(sys.stdout, index=False)

        # (Tùy chọn) Lưu ra file csv hoặc excel mới
        # df_final.to_excel("extracted_data.xlsx", index=False)

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()