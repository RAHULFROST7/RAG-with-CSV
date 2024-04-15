import csv
import json
import pandas as pd
from pathlib import Path

def excel_to_csv(excel_file):
    try:
        df = pd.read_excel(excel_file)
        csv_file = excel_file.parent / f"{excel_file.stem}.csv"
        
        try:
            df = df.drop(['Unnamed: 0'], axis=1)
            df.to_csv(csv_file, index=False)      
        except:
            df.to_csv(csv_file, index=False)
            
        if csv_file.exists():
            print("Converted successfully!")
            return csv_file, excel_file.suffix
        else:
            return "CSV file creation failed."
            
    except Exception as e:
        return f"An error occurred: {str(e)}"


def csv_to_json(csv_file, json_file, url, conversion_ext):
    with open(csv_file, 'r') as f:
        reader = csv.DictReader(f)
        data = []
        for row in reader:
            json_row = {}
            content = ', '.join([f"{key} : {value}" for key, value in row.items()])
            try:
                raw_title = csv_file.name
            except:
                raw_title = str(csv_file).split('\\')[-1]
            doc_name = raw_title.split(".")[0] + conversion_ext
            json_row['title'] = raw_title.split(".")[0]
            json_row['content'] = content
            json_row['url'] = url
            json_row['doc_name'] = doc_name
            
            data.append(json_row)
    
    with open(json_file, 'w') as f:
        json.dump(data, f, indent=4)
        if json_file.exists():
            print(f"\n\nSuccesfully formated data!\nCheck : {json_file} \n\n")
       

def get_data():
    csv_path = Path(input("Please enter a file path: ").strip())
    op_path = Path(input("Please enter your output folder path: ").strip())
    url = input("Please enter file url: ").strip()
    if not url:
        url = "https://drive.google.com/file/d/1yn_5Nauy2QoHhRcdxlMpWH8zsd2Q5REo/view?usp=sharing"
    return csv_path, op_path, url


if __name__ == "__main__":
    
    csv_path, op_path, url = get_data()
    
    if csv_path.suffix == '.csv':
        csv_to_json(csv_path, op_path, url, ".csv")
    
    elif csv_path.suffix == '.xls' or csv_path.suffix == '.xlsx':
        print("Given file is not a CSV")
        csv_path, conversion_flag = excel_to_csv(csv_path)
        csv_to_json(csv_path, op_path, url, conversion_flag)

    else:
        print("Invalid file. Expected a CSV, XLS, or XLSX.")
