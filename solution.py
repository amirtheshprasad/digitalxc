import pandas as pd
import re

# Replace 'your_file.xlsx' with the path to your Excel file
file_path = 'coding_challenge_test.xlsx'

# Read the Excel file
df = pd.read_excel(file_path, usecols=['Comments and Work notes'])

# Display the first few rows of the dataframe
# print(df)
groups = {}
count = 0
# pattern = r'^Groups.*'
# pattern2 = r'^Group Names.*'
pattern3 = r'\[code\]<I>(.*?)</I>\[/code\]'
for index, row in df.iterrows():
    # count = count + 1
    # print(count)
    text = row['Comments and Work notes']
    matches3 = re.findall(pattern3, text, re.MULTILINE)
    # print(matches3)
    for match in matches3:
        if ',' in match:
            # print('yes')
            parts = match.split(',', 1)
            # print(parts)
            for part in parts:
                part = part.strip()
                found = False
                # for group in groups:
                if part in groups:
                    found = True
                    groups[part] = groups[part] + 1
                    break
                if not found:
                    groups[part]=1
        else:
            match = match.strip()
            found = False
            # for group in groups:
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
    
    # matches = re.findall(pattern, text, re.MULTILINE)
    # if matches:
    #     print(matches)
    #     if ',' in text:
    #         matches3 = re.findall(pattern3, text, re.MULTILINE)
    #         print(matches3)
    #     else:
    #         matches3 = re.findall(pattern3, text, re.MULTILINE)
    #         print(matches3)
    # else:
    #     matches2 = re.findall(pattern2, text, re.MULTILINE)
    #     print(matches2)