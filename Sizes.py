import pandas as pd
import re

def extract_size(name):
   
    pattern_size = r'\b(D?\d*\.?\d+(?:x\d*\.?\d+)*(?:H|h|Y|y)?\s*[cC][mM])\b'
    matches_size = re.findall(pattern_size, name)
    size_str = ' '.join(matches_size) if matches_size else None
    return size_str

def move_hm_code_to_end(name):
   
    pattern_hm = r'\bHM\d{4,}\b'
    match_hm = re.search(pattern_hm, name)
    if match_hm:
        hm_code = match_hm.group()
        name_without_hm = re.sub(pattern_hm, '', name).strip()
        name = f"{name_without_hm} {hm_code}"
    return name

def process_excel(input_file, output_file):
   
    df = pd.read_excel(input_file)
    
   
    name_column = df.columns[0]
    
 
    sizes = []
    adjusted_names = []

    for index, row in df.iterrows():
        name = row[name_column]
        size = extract_size(name)
        adjusted_name = move_hm_code_to_end(name)
        sizes.append(size)
        adjusted_names.append(adjusted_name)
        
        print(f"Original name: {name}")
        print(f"Adjusted name: {adjusted_name}")
        print(f"Size: {size}")
        print("---")

    
    df[name_column] = adjusted_names
    df["Size (cm)"] = sizes
    
    
    print(df)
    
    
    df.to_excel(output_file, index=False)


input_file = 'FEED1.xlsx'  
output_file = 'processed_feed1.xlsx'


process_excel(input_file, output_file)
