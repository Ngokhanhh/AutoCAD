import pandas as pd
import argparse
import os
import sys
import json 

def main():
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description="Program to analyze maximum power $P_{max}$ from Excel files.")
    parser.add_argument("input_file", help="Path to the input Excel file")
    args = parser.parse_args()
    file_path = args.input_file

    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        sys.exit(1)

    try:
        df_raw = pd.read_excel(file_path, header=None)
        header_row_index = None
        for i, row in df_raw.iterrows():
            if "TBA" in row.values and "0h:30" in row.values:
                header_row_index = i
                break

        if header_row_index is None:
            return

        df = pd.read_excel(file_path, header=header_row_index)
        target_labels = ['E20.3', 'E22.4', 'E5.7']
        start_col = '0h:30'
        end_col = '24h'

        # --- BIẾN LƯU KẾT QUẢ ĐỂ TRẢ VỀ ---
        structured_results = []

        print(f"\n--- ANALYZING FILE: {os.path.basename(file_path)} ---")

        for tba in target_labels:
            tba_rows = df[df['TBA'] == tba].copy()
            if tba_rows.empty:
                continue

            power_data = tba_rows.loc[:, start_col:end_col].apply(pd.to_numeric, errors='coerce')

            # Nếu trạm có nhiều đường dây, lấy Pmax của dòng TOTAL (Tổng trạm)
            if len(tba_rows) > 1:
                tba_total_series = power_data.sum(axis=0)
                p_max_val = tba_total_series.max()
                t_max_val = tba_total_series.idxmax()
            else:
                # Nếu chỉ có 1 đường dây
                row_p = power_data.iloc[0]
                p_max_val = row_p.max()
                t_max_val = row_p.idxmax()

            # Lưu vào danh sách (đảm bảo kiểu dữ liệu chuẩn để chuyển JSON)
            structured_results.append({
                "station": tba,
                "pmax": float(p_max_val),
                "time": str(t_max_val)
            })
            
            print(f"Trạm {tba}: {p_max_val:.2f} MW lúc {t_max_val}")

        # --- DÒNG QUAN TRỌNG: In JSON để script cha có thể đọc ---
        # Sử dụng tiền tố để script cha dễ nhận diện
        print(f"DATA_JSON:{json.dumps(structured_results)}")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()