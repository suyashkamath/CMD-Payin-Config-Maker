import sys
import pandas as pd
import json

def main():
    if len(sys.argv) != 2:
        print("Usage: python extract_rto.py <excel_file_path>")
        sys.exit(1)

    file_path = sys.argv[1]

    try:
        # Read Excel file
        df = pd.read_excel(file_path)

        # Check if required columns exist
        if 'rto_group_id' not in df.columns or 'rto_group_name' not in df.columns:
            print("Error: Required columns not found in Excel file.")
            sys.exit(1)

        # Extract required columns
        result = []
        for _, row in df.iterrows():
            result.append({
                "id": row["rto_group_id"],
                "name": row["rto_group_name"]
            })

        # Write to JSON file
        with open("output.json", "w", encoding="utf-8") as f:
            json.dump(result, f, indent=4, ensure_ascii=False)

        print("✅ JSON file created successfully: output.json")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
