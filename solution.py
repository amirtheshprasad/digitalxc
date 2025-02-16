import pandas as pd
import re


file_path = 'coding_challenge_test.xlsx'


df = pd.read_excel(file_path, usecols=['Comments and Work notes'])


groups = {}
count = 0

pattern3 = r'\[code\]<I>(.*?)</I>\[/code\]'
for index, row in df.iterrows():
    
    text = row['Comments and Work notes']
    matches3 = re.findall(pattern3, text, re.MULTILINE)

    for match in matches3:
        if ',' in match:

            parts = match.split(',', 1)

            for part in parts:
                part = part.strip()
                found = False

                if part in groups:
                    found = True
                    groups[part] = groups[part] + 1
                    break
                if not found:
                    groups[part]=1
        else:
            match = match.strip()
            found = False
            
            if match in groups:
                found = True
                groups[match] = groups[match] + 1
                break
            if not found:
                groups[match]=1
    
    
print(groups)
df_groups = pd.DataFrame(list(groups.items()), columns=['Group name', 'Number of occurrences'])
output_file_path = 'output.xlsx'
df_groups.to_excel(output_file_path, index=False)

print("Output written to", output_file_path)
    
    