import subprocess
import json
import sys
import pathlib
import argparse
import re
import os

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

def scan_and_analyze(folder_path):
    excel_files = get_excel_file_paths(folder_path)
    if not excel_files:
        print("No matching files found to process.")
        return

    # Dictionary to store the highest Pmax records
    monthly_summary = {} 

    print(f"Starting process for {len(excel_files)} files...")

    # Force UTF-8 environment for subprocesses
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    for file_path in excel_files:
        try:
            result = subprocess.run(
                [sys.executable, "Extract_Data_From_Excel_v2.py", file_path],
                capture_output=True, 
                text=True, 
                encoding='utf-8',
                env=env,
                errors='replace'
            )

            # 1. Log detailed text output to Summary.txt
            with open("Summary.txt", "a", encoding="utf-8") as f:
                f.write(f"\n--- FILE: {pathlib.Path(file_path).name} ---\n")
                f.write(result.stdout)
                if result.stderr:
                    f.write(f"STDERR: {result.stderr}\n")

            # 2. Parse JSON from stdout to update high scores
            for line in result.stdout.splitlines():
                if line.startswith("DATA_JSON:"):
                    json_str = line.replace("DATA_JSON:", "").strip()
                    data_list = json.loads(json_str)
                    
                    for item in data_list:
                        station = item['station']
                        pmax = item['pmax']
                        time = item['time']
                        
                        # Comparison logic for maximum values
                        if station not in monthly_summary or pmax > monthly_summary[station]['pmax']:
                            monthly_summary[station] = {
                                'pmax': pmax,
                                'time': time,
                                'source_file': pathlib.Path(file_path).name
                            }
                
        except Exception as e:
            print(f"Error processing file {file_path}: {e}")

    # --- ACTION: SAVE FINAL AGGREGATED DATA TO JSON FILE ---
    output_json_file = "monthly_summary.json"
    try:
        with open(output_json_file, "w", encoding="utf-8") as jf:
            json.dump(monthly_summary, jf, indent=4, ensure_ascii=False)
        print(f"\nFinal monthly summary successfully saved to: {output_json_file}")
    except Exception as e:
        print(f"Error saving JSON report: {e}")

    # --- PRINT FINAL REPORT TO TERMINAL ---
    print("\n" + "="*85)
    print(f"{'SUMMARY REPORT':^85}")
    print("="*85)
    print(f"{'Substation':<15} | {'Pmax (MW)':<12} | {'Max Time':<12} | {'Source File'}")
    print("-"*85)
    
    for st, val in monthly_summary.items():
        print(f"{st:<15} | {val['pmax']:>12.2f} | {val['time']:<12} | {val['source_file']}")
    print("="*85)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scan folders and export peak monthly $P_{max}$ to JSON.")
    parser.add_argument("folder", nargs='+', help="Root folder path to scan")
    
    args = parser.parse_args()
    full_folder_path = " ".join(args.folder)
    scan_and_analyze(full_folder_path)