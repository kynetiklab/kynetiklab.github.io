```python
#DOWNLOADING OF FILES FROM THE WEBSITE AND SAVING INSIDE FOLDER "DATA INSURANCE"

import os
import requests
import pandas as pd

# Base URL of the files
base_url = "https://www.ia.org.hk/en/infocenter/statistics/files/{}q{}long.xls"

# Specify the path where files will be saved
save_folder = r"C:\Users\User\Desktop\Data Insurance"

# Create the save folder if it doesn't exist
os.makedirs(save_folder, exist_ok=True)

# Loop through years from 2001 to 2023
for year in range(2014, 2024):
    for quarter in range(1, 5):  # Quarters: Q1 to Q4
        file_url = base_url.format(quarter, str(year)[2:])
        file_path = os.path.join(save_folder, f"{year} Q{quarter}.xls")
        
        response = requests.get(file_url)
        
        if response.status_code == 200:
            with open(file_path, "wb") as file:
                file.write(response.content)
            print(f"File {year}_Q{quarter}.xls downloaded successfully.")
            
            # Read the downloaded Excel file into a DataFrame
            df = pd.read_excel(file_path, skiprows=2)

            # Define the clean_data function
            def clean_data(df):
                excluded_cols = ["Name of Insurer"]
                df_cleaned = df.copy()
                
                for col_name in df.columns:
                    if col_name not in excluded_cols:
                        df_cleaned[col_name] = df[col_name].replace("-", "0")
                
                df_cleaned = df_cleaned.replace("-", pd.NA)
                df_cleaned = df_cleaned.fillna(0)
                
                return df_cleaned

            # Clean the DataFrame
            cleaned_df = clean_data(df)

            # Perform data analysis on the cleaned DataFrame
            # ...
            
        else:
            print(f"Failed to download {year}_Q{quarter}.xls. Status code: {response.status_code}")

```


```python
#CLEANING, REFORMATTING, AND SAVING INTO FOLDER "OUTPUT"

import pandas as pd
import glob
import os
import time

input_directory = r'C:\Users\User\Desktop\Data Insurance'  # Replace with your input directory path
output_directory = r'C:\Users\User\Desktop\Data Insurance\output'  # Replace with your output directory path

# Define the new header and columns to replace '-' with 0 for Table L1
new_header = [
    "Name of Insurer",
    "Non-Linked(Class A): Life & Annuity(All Coverages) Single Revenue Premiums (thousands of HKD)",
    "Non-Linked(Class A): Life & Annuity (All Coverages) Annualized Premiums (thousands of HKD)",
    "Non-Linked(Class A): Accident & Sickness (Medical Part Only) Single Revenue Premiums (thousands of HKD)",
    "Non-Linked(Class A): Accident & Sickness (Medical Part Only) Annualized Premiums (thousands of HKD)",
    "Linked(Class C): Life & Annuity (All Coverages) Single Revenue Premiums (thousands of HKD)",
    "Linked(Class C): Life & Annuity (All Coverages) Annualized Premiums (thousands of HKD)",
    "Linked(Class C): Accident & Sickness(Medical Part Only) Single Revenue Premiums (thousands of HKD)",
    "Linked(Class C): Accident & Sickness(Medical Part Only) Annualized Premiums (thousands of HKD)",
    "Linked/Non-Linked: Others(Classes B, D, E, F): Single Revenue Premiums (thousands of HKD)",
    "Linked/Non-Linked: Others(Classes B, D, E, F):  Annualized Premiums (thousands of HKD)",
    "Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)",
    "Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)"
]

# Define the new header and columns to replace '-' with 0 for Table L1(a)
new_header_l1a = [
    "Name of Insurer",
    "Currency: Policy Issued in HKD Single Revenue Premiums (thousands of HKD)",
    "Currency: Policy Issued in HKD Annualized Premiums (thousands of HKD)",
    "Currency: Policy Issued in RMB Single Revenue Premiums (thousands of HKD)",
    "Currency: Policy Issued in RMB Annualized Premiums (thousands of HKD)",
    "Currency: Policy Issued in USD Single Revenue Premiums (thousands of HKD)",
    "Currency: Policy Issued in USD Annualized Premiums (thousands of HKD)",
    "Currency: Policy Issued in Other Currencies Single Revenue Premiums (thousands of HKD)",
    "Currency: Policy Issued in Other Currencies Annualized Premiums (thousands of HKD)",
    "Currency: Total Single Revenue Premiums (thousands of HKD)",
    "Currency: Total Annualized Premiums (thousands of HKD)"
]

# Define the new header and columns for Table L1(b)
new_header_l1b = [
    "Name of Insurer",
    "Onshore / Offshore: Onshore Single Revenue Premiums (thousands of HKD)",
    "Onshore / Offshore: Onshore Annualized Premiums (thousands of HKD)",
    "Onshore / Offshore: Offshore Single Revenue Premiums (thousands of HKD)",
    "Onshore / Offshore: Offshore Annualized Premiums (thousands of HKD)",
    "Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)",
    "Onshore / Offshore: Total Annualized Premiums (thousands of HKD)"
]
# Define the new header and columns for Table L1(c)
new_header_l1c = [
    "Name of Insurer",
    " Premium Term: Single Revenue Premiums (thousands of HKD)",
    "Premium Term: Annualized Premiums(<5 years) (thousands of HKD)",
    "Premium Term: Annualized Premiums  (5 <10 years) (thousands of HKD)",
    "Premium Term: Annualized Premiums (10 <25 years) (thousands of HKD)",
    "Premium Term: Annualized Premiums (25+ years) (thousands of HKD)",
    "Premium Term: Total of Annualized Premiums: (thousands of HKD)"
]

# Define the new header and columns for Table L1(d)
new_header_l1d = [
    "Name of Insurer",
    "Distribution Channel: Agents(Excluding Banks)Single Revenue Premiums (thousands of HKD)",
    "Distribution Channel: Agents(Excluding Banks) Annualized Premiums (thousands of HKD)",
    "Distribution Channel: Banks      Single Revenue Premiums (thousands of HKD)",
    "Distribution Channel: Banks     Annualized Premiums (thousands of HKD)",
    "Distribution Channel: Brokers      Single Revenue Premiums (thousands of HKD)",
    "Distribution Channel: Brokers    Annualized Premiums (thousands of HKD)",
    "Distribution Channel: Direct      Single Revenue Premiums (thousands of HKD)",
    "Distribution Channel: Direct    Annualized Premiums (thousands of HKD)",
    "Distribution Channel: Others      Single Revenue Premiums (thousands of HKD)",
    "Distribution Channel: Others     Annualized Premiums (thousands of HKD)",
    "Channel Distribution: Total Single Revenue Premiums (thousands of HKD)",
    "Channel Distribution: Total Annualized Premiums (thousands of HKD)"
]
# Define the new header and columns to replace '-' with 0 for Table L1(e)
new_header_l1e = [
    "Name of Insurer",
    "Policy Issued in HKD: Number of Policies; Single Premiums",
    "Policy Issued in HKD: Number of Policies; Non-Single Premiums",
    "Policy Issued in RMB: Number of Policies; Single Premiums",
    "Policy Issued in RMB: Number of Policies; Non-Single Premiums",
    "Policy Issued in USD: Number of Policies; Single Premiums",
    "Policy Issued in USD: Number of Policies; Non-Single Premiums",
    "Policy Issued in Other Currencies: Number of Policies; Single Premiums",
    "Policy Issued in Other Currencies: Number of Policies; Non-Single Premiums",
    "Total: Number of Policies; Single Premiums",
    "Total: Number of Policies; Non-Single Premiums"
]
# Define the new header and columns to replace '-' with 0 for Table L1(f)
new_header_l1f = [
    "Name of Insurer",
    "Onshore / Offshore: Onshore Number of Policies Single Premiums",
    "Onshore / Offshore: Onshore Number of Policies Non-Single Premiums",
    "Onshore / Offshore: Offshore Number of Policies Single Premiums",
    "Onshore / Offshore: Offshore Number of Policies Non-Single Premiums",
    "Onshore / Offshore: Total Onshore Number of Policies Single Revenue Premiums",
    "Onshore / Offshore: Total Offshore Number of Policies Non-Single Premiums"
]
# Define the new header and columns to replace '-' with 0 for Table L1(g)
new_header_l1g = [
    "Name of Insurer",
    "Premium Term: Number of Policies; Single Premiums",
    "Premium Term: Number of Policies; Non-Single Premiums (<5 years) (thousands of HKD)",
    "Premium Term: Number of Policies; Non-Single Premiums (5 <10 years) (thousands of HKD)",
    "Premium Term: Number of Policies; Non-Single Premiums (10 <25 years) (thousands of HKD)",
    "Premium Term: Number of Policies; Non-Single Premiums (25+ years) (thousands of HKD)",
    "Premium Term: Number of Policies; Total of Non-Single Premiums: (thousands of HKD)"
]
# Define the new header and columns to replace '-' with 0 for Table L1(h)
new_header_l1h = [
    "Name of Insurer",
    "Distribution Channel: Agents(Excluding Banks) Number of Policies Single Premiums",
    "Distribution Channel: Agents(Excluding Banks) Number of Policies Non-Single Premiums",
    "Distribution Channel: Banks Number of Policies Single Premiums",
    "Distribution Channel: Banks Number of Policies Non-Single Premiums",
    "Distribution Channel: Brokers Number of Policies Single Premiums",
    "Distribution Channel: Brokers Number of Policies Non-Single Premiums",
    "Distribution Channel: Direct Number of Policies Single Premiums",
    "Distribution Channel: Direct Number of Policies Non-Single Premiums",
    "Distribution Channel: Others Number of Policies Single Premiums",
    "Distribution Channel: Others Number of Policies Non-Single Premiums",
    "Channel Distribution:Total Number of Policies Single Premiums",
    "Channel Distribution: Total Number of Policies Non-Single Premiums"
]
# Define the new header and columns to replace '-' with 0 for Table L2
new_header_l2 = [
    "Name of Insurer",
    "Number of Policies",
    "Number of Lives",
    "Single Revenue Premiums",
    "Annualized Premiums"
]
# Define the new header for consolidated Table L3-1
new_header_l31 = [
    "Name of Insurer",
    "Non-Linked Individual Business (Class A): Number of Policies",
    "Non-Linked Individual Business (Class A): Sums Assured (thousands of HKD)",
    "Non-Linked Individual Business (Class A): Single Revenue Premiums (thousands of HKD)",
    "Non-Linked Individual Business (Class A): Non-Single Revenue Premiums (HK$'000",
    "Linked Individual Business (Class C): Number of Policies",
    "Linked Individual Business (Class C): Sums Assured (thousands of HKD)",
    "Linked Individual Business (Class C): Single Revenue Premiums (thousands of HKD)",
    "Linked Individual Business (Class C): Non-Single Revenue Premiums (thousands of HKD)",
]  
# Define the new header for consolidated Table L3-2
new_header_l32 = [ 
    "Name of Insurer",
    "Other Individual Business (Classes B, D, E & F): Number of Policies",
    "Other Individual Business (Classes B, D, E & F): Single Revenue (thousands of HKD)",
    "Other Individual Business (Classes B, D, E & F): Non-Single Revenue (thousands of HKD)",
    "Total Individual Business: Number of Policies",
    "Total Individual Business: Sums Assured (thousands of HKD)",
    "Total Individual Business: Single Revenue (thousands of HKD)",
    "Total Individual Business: Non-Single Revenue"
]
# Define the new header for consolidated Table L4
new_header_l4 = [ 
    "Name of Insurer",
    "Non-Retirement Scheme Group Business (Classes A to F & I): Number of Policies",
    "Non-Retirement Scheme Group Business (Classes A to F & I): Number of Lives",
    "Non-Retirement Scheme Group Business (Classes A to F & I): Single Revenue Premiums (thousands of HKD)",
    "Non-Retirement Scheme Group Business (Classes A to F & I): Non-Single Revenue Premiums (thousands of HKD)",
    "Retirement Scheme Group Business (Classes G & H): Number of Schemes",
    "Retirement Scheme Group Business (Classes G & H): Ending Fund Balance (thousands of HKD)",
    "Retirement Scheme Group Business (Classes G & H): Single Contributions (thousands of HKD)",
    "Retirement Scheme Group Business (Classes G & H): Non-Single Revenue Contributions (thousands of HKD)"
]

# Function to replace '-' with 0 in a DataFrame
def replace_dash_with_zero(df):
    return df.applymap(lambda x: 0 if x == '-' else x)

# Iterate through each Excel file in the input directory
for input_file in glob.glob(os.path.join(input_directory, '*.xls')):
    xls = pd.ExcelFile(input_file)
    
    for sheet_name in xls.sheet_names:
        if sheet_name == 'Table L1':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1
            df.columns = new_header
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Define the output file path
            output_filename = os.path.splitext(os.path.basename(input_file))[0] + '.xlsx'
            output_path = os.path.join(output_directory, output_filename)
            
            # Save the reformatted DataFrame to a new Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Table L1', index=False)
            time.sleep(2)
            
        elif sheet_name == 'Table L1(a)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(a)
            df.columns = new_header_l1a
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1a[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1a[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(a)', index=False)
            time.sleep(2)
        
        elif sheet_name == 'Table L1(b)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(b)
            df.columns = new_header_l1b
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1b[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1b[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(b)', index=False)
            time.sleep(2)
            
        elif sheet_name == 'Table L1(c)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(c)
            df.columns = new_header_l1c
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[12, "Name of Insurer"]
            numeric_values = df.loc[12, new_header_l1c[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1c[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:13], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(c)', index=False)
            time.sleep(2)
            
        elif sheet_name == 'Table L1(d)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(d)
            df.columns = new_header_l1d
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1d[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1d[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(d)', index=False)
            time.sleep(2)
                
        elif sheet_name == 'Table L1(e)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(e)
            df.columns = new_header_l1e
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1e[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1e[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(e)', index=False)
            time.sleep(2)
                
        elif sheet_name == 'Table L1(f)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(f)
            df.columns = new_header_l1f
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1f[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1f[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(f)', index=False)  
            time.sleep(2)
                
        elif sheet_name == 'Table L1(g)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(g)
            df.columns = new_header_l1g
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1g[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1g[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(g)', index=False)
            time.sleep(2)
                
        elif sheet_name == 'Table L1(h)':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(h)
            df.columns = new_header_l1h
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[13, "Name of Insurer"]
            numeric_values = df.loc[13, new_header_l1h[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l1h[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:14], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L1(h)', index=False)
            time.sleep(2)
                
        elif sheet_name == 'Table L2':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L2
            df.columns = new_header_l2
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[9, "Name of Insurer"]
            numeric_values = df.loc[9, new_header_l2[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l2[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:10], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L2', index=False)
            time.sleep(2)
                
        elif sheet_name == 'Table L3-1':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(g)
            df.columns = new_header_l31
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[11, "Name of Insurer"]
            numeric_values = df.loc[11, new_header_l31[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l31[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:12], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L3-1', index=False)
            time.sleep(2)
                
        elif sheet_name == 'Table L3-2':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L1(h)
            df.columns = new_header_l32
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[11, "Name of Insurer"]
            numeric_values = df.loc[11, new_header_l32[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l32[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:12], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L3-2', index=False)
            time.sleep(2)
        elif sheet_name == 'Table L4':
            df = pd.read_excel(xls, sheet_name, header=None)
            
            # Remove the second column (column B)
            df.drop(columns=[1], inplace=True)
            
            # Set the new header for Table L4
            df.columns = new_header_l4
            
            # Extract values from specific cells and update the DataFrame
            name_of_insurer = df.loc[12, "Name of Insurer"]
            numeric_values = df.loc[12, new_header_l4[1:]]
            
            df.loc[0, "Name of Insurer"] = name_of_insurer
            df.loc[0, new_header_l4[1:]] = numeric_values
            
            # Drop unnecessary rows and reset index
            df.drop(df.index[1:13], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            # Replace '-' with 0 in numeric cells
            df = replace_dash_with_zero(df)
            
            # Append the reformatted DataFrame to the existing Excel file in .xlsx format
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, sheet_name='Table L4', index=False)
            time.sleep(2)
        
        
    print(f"Processed: {input_file}")

```


```python
#TABLE L3-1 AND L3-3 COMBINATION
import os
from openpyxl import load_workbook

output_folder = r'C:\Users\User\Desktop\Data Insurance\output'
col_index_of_insurer_column = 0  # Adjust to the actual column index of the insurer name column

def combine_sheets_horizontally(sheet1, sheet2, target_sheet):
    for row1, row2 in zip(sheet1.iter_rows(min_row=2, values_only=True), sheet2.iter_rows(min_row=2, values_only=True)):
        combined_row = list(row1) + list(row2)[1:]  # Exclude the first column of sheet2
        target_sheet.append(combined_row)

for filename in os.listdir(output_folder):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(output_folder, filename)
        wb = load_workbook(file_path)
        
        if 'Table L3-1' in wb.sheetnames and 'Table L3-2' in wb.sheetnames:
            sheet_l3_1 = wb['Table L3-1']
            sheet_l3_2 = wb['Table L3-2']
            
            # Create or get the combined sheet
            combined_sheet = wb.create_sheet('Table L3')
            
            # Append header row to combined sheet
            header_row = [sheet_l3_1.cell(row=1, column=col_index_of_insurer_column + 1).value]
            for cell in sheet_l3_1.iter_cols(min_col=2, max_col=sheet_l3_1.max_column):
                header_row.append(cell[0].value)
            for cell in sheet_l3_2.iter_cols(min_col=2, max_col=sheet_l3_2.max_column):
                header_row.append(cell[0].value)
            combined_sheet.append(header_row)
            
            combine_sheets_horizontally(sheet_l3_1, sheet_l3_2, combined_sheet)
            
            wb.save(file_path)
            
            print(f'Combination complete for {filename}')

```


```python
#ANALYSIS AND USER INTERFACE SCRIPT
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

def load_insurance_data(data_directory, start_year, start_quarter, end_year, end_quarter):
    sheet_names_L1 = ['Table L1']
    sheet_names_L1_b = ['Table L1(b)']
    sheet_names_L1_d = ['Table L1(d)']
    df_dict_L1 = {}  # Store DataFrames for Table L1 sheets in a dictionary
    df_dict_L1_b = {}  # Store DataFrames for Table L1(b) sheets in a dictionary
    df_dict_L1_d = {}  # Store DataFrames for Table L1(d) sheets in a dictionary

    # Get a list of all insurance files in the directory
    insurance_files = [f for f in os.listdir(data_directory) if f.endswith('.xlsx') and not f.startswith('~$')]

    for insurance_file in insurance_files:
        # Extract the year and quarter from the file name
        year, quarter = extract_year_quarter(insurance_file)

        # Check if the file's year-quarter is within the desired range
        if start_year <= year <= end_year:
            if start_year == end_year:  # Case: Single year
                if start_quarter <= quarter <= end_quarter:
                    load_data_into_dict(insurance_file, data_directory, sheet_names_L1, sheet_names_L1_b, sheet_names_L1_d, df_dict_L1, df_dict_L1_b, df_dict_L1_d)
            else:  # Case: Multiple years
                if year == start_year and start_quarter <= quarter:
                    load_data_into_dict(insurance_file, data_directory, sheet_names_L1, sheet_names_L1_b, sheet_names_L1_d, df_dict_L1, df_dict_L1_b, df_dict_L1_d)
                elif year == end_year and quarter <= end_quarter:
                    load_data_into_dict(insurance_file, data_directory, sheet_names_L1, sheet_names_L1_b, sheet_names_L1_d, df_dict_L1, df_dict_L1_b, df_dict_L1_d)
                elif start_year < year < end_year:
                    load_data_into_dict(insurance_file, data_directory, sheet_names_L1, sheet_names_L1_b, sheet_names_L1_d, df_dict_L1, df_dict_L1_b, df_dict_L1_d)

    return df_dict_L1, df_dict_L1_b, df_dict_L1_d

def load_data_into_dict(insurance_file, data_directory, sheet_names_L1, sheet_names_L1_b, sheet_names_L1_d, df_dict_L1, df_dict_L1_b, df_dict_L1_d):
    file_path = os.path.join(data_directory, insurance_file)

    for sheet_name in sheet_names_L1:
        try:
            df_L1 = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            df_dict_L1[(insurance_file, sheet_name)] = df_L1
        except Exception as e:
            print(f"Error loading data from sheet '{sheet_name}' in file '{insurance_file}': {e}")

    for sheet_name_b in sheet_names_L1_b:
        try:
            df_L1_b = pd.read_excel(file_path, sheet_name=sheet_name_b, engine='openpyxl')
            df_dict_L1_b[(insurance_file, sheet_name_b)] = df_L1_b
        except Exception as e:
            print(f"Error loading data from sheet '{sheet_name_b}' in file '{insurance_file}': {e}")

    for sheet_name_d in sheet_names_L1_d:
        try:
            df_L1_d = pd.read_excel(file_path, sheet_name=sheet_name_d, engine='openpyxl')
            df_dict_L1_d[(insurance_file, sheet_name_d)] = df_L1_d
        except Exception as e:
            print(f"Error loading data from sheet '{sheet_name_d}' in file '{insurance_file}': {e}")

def extract_year_quarter(file_name):
    # Assumes the file name is in the format: 'yyyy Qx.xlsx' where x is the quarter number
    parts = file_name.split()
    if len(parts) == 2 and parts[0].isdigit() and parts[1].startswith('Q') and parts[1][1:-5].isdigit():
        year_str, quarter_str = parts[0], parts[1][1:-5]
        year = int(year_str)
        quarter = int(quarter_str)
        return year, quarter
    else:
        # Handle other filename formats or invalid filenames
        raise ValueError(f"Invalid filename format: {file_name}")



# Function to format y-axis labels as thousands
def format_thousands(x, pos):
    'The two args are the value and tick position'
    return f'{int(x/1000):,}K'

# Total Revenue for Top 20 Insurers from Table L1
def Total_Revenue_Comparison_For_Top_20_Insurers_L1(df_dict_L1):
    # Loop through the loaded DataFrames and perform analysis for each file and sheet
    for (file, sheet_name), df in df_dict_L1.items():
        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        df['Total Revenue'] = (df['Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)']

        # Exclude the 'Market Total' row if it exists in the DataFrame
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer and sort in descending order
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum().sort_values(ascending=False)

        # Get the top 20 insurers
        top_20_insurers = total_revenue_by_insurer.head(20)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print(f"Top 20 Insurers - Total Revenue (thousands of HKD):")
        print(top_20_insurers)

        # Create a bar chart to compare the Total Revenue for the top 20 insurers
        plt.figure(figsize=(12, 6))
        ax = top_20_insurers.plot(kind='bar')
        plt.xlabel('Insurer')
        plt.ylabel('Total Revenue (thousands of HKD)')
        plt.title(f'Top 20 Insurers - Total Revenue for {file} - Sheet: {sheet_name}')

        # Format y-axis labels to display actual figures in thousands (K)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(format_thousands))

        plt.show()
       
# Single Revenue Vs Annualized Table L1
def Single_Vs_Annualized_Comparison_Bar_Chart_L1(df_dict_L1):
    # Loop through the loaded DataFrames and perform analysis for each file and sheet
    for (file, sheet_name), df in df_dict_L1.items():
        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Exclude the 'Market Total' row if it exists in the DataFrame
        df = df[df['Name of Insurer'] != 'Market Total']

        # Find the columns containing 'Linked/Non-Linked: Total Single Revenue Premiums' and 'Linked/Non-Linked: Total Annualized Premiums'
        total_single_revenue_column = df.columns[df.columns.str.contains('Linked/Non-Linked:\s*Total\s*Single\s*Revenue\s*Premiums', case=False, na=False)].tolist()[0]
        total_annualized_premiums_column = df.columns[df.columns.str.contains('Linked/Non-Linked:\s*Total\s*Annualized\s*Premiums', case=False, na=False)].tolist()[0]

        # Calculate the total single revenue premiums and total annualized premiums for each insurer
        df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]

        # Sort the DataFrame by 'Total Revenue' column in descending order
        df = df.sort_values(by='Total Revenue', ascending=False)

        # Get the top 20 insurers
        top_20_insurers = df.head(20)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print(f"Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):")
        print(top_20_insurers[['Name of Insurer', 'Total Revenue']])

        # Create a bar chart to compare the Total Revenue for the top 20 insurers
        plt.figure(figsize=(12, 6))
        ax = plt.gca()
        x = range(len(top_20_insurers))
        ax.bar(x, top_20_insurers['Total Revenue'], label='Total Revenue')
        plt.xlabel('Insurer')
        plt.ylabel('Total Revenue (thousands of HKD)')
        plt.title(f'Top 20 Insurers - Linked/Non-Linked: Total Revenue for {file} - Sheet: {sheet_name}')
        plt.xticks(x, top_20_insurers['Name of Insurer'], rotation=90)
        plt.legend()

        # Format y-axis labels to display actual figures in thousands (K)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(format_thousands))

        plt.show()
        
# Market Share Totals for Table L1
def Market_Share_Totals_L1(df_dict_L1):
    for key, df in df_dict_L1.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        total_revenue_formula = (df['Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print("Total Revenue by Insurer:")
        print(total_revenue_by_insurer)
        print("\nMarket Share by Insurer (%):")
        print(market_share_by_insurer)
        print("\n")

       
# Market Share Top 10 Analysis Pie Chart for Table L1
def Market_Share_Top_10_Pie_Chart_L1(df_dict_L1):
    for key, df in df_dict_L1.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        df['Total Revenue'] = (df['Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)']

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Separate the top 10 insurers and combine the rest as "Others"
        top_10_insurers = market_share_by_insurer.head(10)
        others_market_share = market_share_by_insurer.iloc[10:].sum()
        top_10_insurers['Others'] = others_market_share

        # Plotting the pie chart
        plt.figure(figsize=(10, 6))
        plt.pie(top_10_insurers, labels=top_10_insurers.index, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
        plt.axis('equal')
        plt.title(f"Market Share Analysis - {file} - {sheet_name}")

        # Show the plot
        plt.show()
        
# Market Share Top 10 Analysis Horizontal Bar Chart for Table L1
def Market_Share_Top_10_Horizontal_Bar_Chart_L1(df_dict_L1):
    for key, df in df_dict_L1.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        total_revenue_formula = (df['Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Separate the top 20 insurers and combine the rest as "Others"
        top_20_insurers = market_share_by_insurer.head(20)
        others_market_share = market_share_by_insurer.iloc[20:].sum()
        top_20_insurers['Others'] = others_market_share

        # Plotting the horizontal bar chart
        plt.figure(figsize=(10, 6))
        top_20_insurers.plot(kind='barh', color='skyblue', edgecolor='black')
        plt.xlabel('Market Share (%)')
        plt.ylabel('Insurer')
        plt.title(f"Market Share Analysis - {file} - {sheet_name}")

        # Show the plot
        plt.show()
        

#Classes Comparison Bar Chart Table L1        
def Classes_Comparisons_L1_Bar_Chart(df_dict_L1):
    for key, df in df_dict_L1.items():
        file, sheet_name = key

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Find the columns containing 'Premiums' and 'Revenues' (ignoring case and whitespace)
        premiums_columns = df.columns[df.columns.str.contains('Premiums', case=False, na=False)].tolist()
        revenues_columns = df.columns[df.columns.str.contains('Revenues', case=False, na=False)].tolist()

        # Drop the unwanted columns
        unwanted_columns = ['Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)',
                            'Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)']
        premiums_columns = [col for col in premiums_columns if col not in unwanted_columns]
        revenues_columns = [col for col in revenues_columns if col not in unwanted_columns]

        # Calculate the total revenue for each insurer based on the formula
        df['Total Revenue'] = df['Linked/Non-Linked: Total Single Revenue Premiums (thousands of HKD)'] / 10 + df['Linked/Non-Linked: Total Annualized Premiums (thousands of HKD)']

        # Calculate the total revenue for each insurer and sort the insurers based on their total revenue
        total_revenue_column = 'Total Revenue'
        total_revenue_by_insurer = df.groupby('Name of Insurer')[total_revenue_column].sum()
        sorted_top_ten_insurers = total_revenue_by_insurer.sort_values(ascending=False).head(10).index.tolist()

        # Select only the relevant columns for the Premium and Revenue Comparison and filter top ten insurers
        relevant_columns = ['Name of Insurer'] + premiums_columns + revenues_columns + [total_revenue_column]
        df = df[relevant_columns]
        df_top_ten = df[df['Name of Insurer'].isin(sorted_top_ten_insurers)].copy()

        # Plotting the stacked bar chart for top ten insurers
        plt.figure(figsize=(12, 8))
        ax = df_top_ten.set_index('Name of Insurer')[premiums_columns + revenues_columns + [total_revenue_column]].plot(kind='bar', stacked=True)
        plt.xlabel('Insurer')
        plt.ylabel('Amount (thousands of HKD)')
        plt.title(f"Linked, Non-Linked & Others - {file} - {sheet_name}")

        # Format the y-axis ticks with thousands separator
        ax.get_yaxis().set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

        # Move the legend outside the plot area to avoid blocking
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1))

        # Show the plot
        plt.show()
        
# Total Revenue for Top 20 Insurers from Table L1(b)
def Total_Revenue_Comparison_For_Top_20_Insurers_L1b(df_dict_L1_b):
    # Loop through the loaded DataFrames and perform analysis for each file and sheet
    for (file, sheet_name), df in df_dict_L1_b.items():
        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        df['Total Revenue'] = (df['Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Onshore / Offshore: Total Annualized Premiums (thousands of HKD)']

        # Exclude the 'Market Total' row if it exists in the DataFrame
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer and sort in descending order
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum().sort_values(ascending=False)

        # Get the top 20 insurers
        top_20_insurers = total_revenue_by_insurer.head(20)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print(f"Top 20 Insurers - Total Revenue (thousands of HKD):")
        print(top_20_insurers)

        # Create a bar chart to compare the Total Revenue for the top 20 insurers
        plt.figure(figsize=(12, 6))
        ax = top_20_insurers.plot(kind='bar')
        plt.xlabel('Insurer')
        plt.ylabel('Total Revenue (thousands of HKD)')
        plt.title(f'Top 20 Insurers - Total Revenue for {file} - Sheet: {sheet_name}')

        # Format y-axis labels to display actual figures in thousands (K)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(format_thousands))

        plt.show()

        
#Market Share Totals

def Market_Share_Totals_L1b(df_dict_L1_b):
    for key, df in df_dict_L1_b.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        df['Total Revenue'] = (df['Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Onshore / Offshore: Total Annualized Premiums (thousands of HKD)']

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print("Total Revenue by Insurer:")
        print(total_revenue_by_insurer)
        print("\nMarket Share by Insurer (%):")
        print(market_share_by_insurer)
        print("\n")

       
def Onshore_Offshore_Market_Share_Top_20_Insurers_Pie_Chart_L1b(df_dict_L1_b):
    for key, df in df_dict_L1_b.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        df['Total Revenue'] = (df['Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Onshore / Offshore: Total Annualized Premiums (thousands of HKD)']

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Separate the top 10 insurers and combine the rest as "Others"
        top_10_insurers = market_share_by_insurer.head(10)
        others_market_share = market_share_by_insurer.iloc[10:].sum()
        top_10_insurers['Others'] = others_market_share

        # Plotting the pie chart
        plt.figure(figsize=(10, 6))
        plt.pie(top_10_insurers, labels=top_10_insurers.index, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
        plt.axis('equal')
        plt.title(f"Onshore/Offshore Market Share Analysis - {file} - {sheet_name}")

        # Show the plot
        plt.show()

        
def Onshore_Offshore_Market_Share_Top_20_Insurers_Horizontal_Bar_L1b(df_dict_L1_b):
    for key, df in df_dict_L1_b.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        df['Total Revenue'] = (df['Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Onshore / Offshore: Total Annualized Premiums (thousands of HKD)']

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Separate the top 20 insurers and combine the rest as "Others"
        top_20_insurers = market_share_by_insurer.head(20)
        others_market_share = market_share_by_insurer.iloc[20:].sum()
        top_20_insurers['Others'] = others_market_share

        # Plotting the horizontal bar chart
        plt.figure(figsize=(10, 6))
        top_20_insurers.plot(kind='barh', color='skyblue', edgecolor='black')
        plt.xlabel('Market Share (%)')
        plt.ylabel('Insurer')
        plt.title(f"Onshore/Offshore Market Share Analysis - {file} - {sheet_name}")

        # Show the plot
        plt.show()

def Premium_Comparison_For_Top_10_Insurers_L1b(df_dict_L1_b):
    for key, df in df_dict_L1_b.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on formula and add it as a new column
        total_revenue_formula = (df['Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Onshore / Offshore: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        
        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Find the columns containing 'Premiums' and 'Revenues' (ignoring case and whitespace)
        premiums_columns = df.columns[df.columns.str.contains('Premiums', case=False, na=False)].tolist()
        revenues_columns = df.columns[df.columns.str.contains('Revenues', case=False, na=False)].tolist()

        # Drop the unwanted columns
        unwanted_columns = [
            'Onshore / Offshore: Total Single Revenue Premiums (thousands of HKD)',
            'Onshore / Offshore: Total Annualized Premiums (thousands of HKD)',
            'Total Revenue (thousands of HKD)'  # Adding 'Total Revenue' to the unwanted list
        ]
        premiums_columns = [col for col in premiums_columns if col not in unwanted_columns]
        revenues_columns = [col for col in revenues_columns if col not in unwanted_columns]

        # Calculate the total revenue for each insurer and sort the insurers based on their total revenue
        total_revenue_column = 'Total Revenue'
        total_revenue_by_insurer = df.groupby('Name of Insurer')[total_revenue_column].sum()
        sorted_top_ten_insurers = total_revenue_by_insurer.sort_values(ascending=False).head(10).index.tolist()

        # Select only the relevant columns for the Premium and Revenue Comparison and filter top ten insurers
        relevant_columns = ['Name of Insurer'] + premiums_columns + revenues_columns + [total_revenue_column]
        df_top_ten = df[df['Name of Insurer'].isin(sorted_top_ten_insurers)][relevant_columns].copy()

        # Plotting the stacked bar chart for top ten insurers
        plt.figure(figsize=(12, 8))
        ax = df_top_ten.set_index('Name of Insurer')[premiums_columns + revenues_columns].plot(kind='bar', stacked=True)
        plt.xlabel('Insurer')
        plt.ylabel('Amount (thousands of HKD)')
        plt.title(f"Premium and Revenue Comparison for Top Ten Insurers - {file} - {sheet_name}")

        # Format the y-axis ticks with thousands separator
        ax.get_yaxis().set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

        # Move the legend outside the plot area to avoid blocking
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1))

        # Show the plot
        plt.show()

#Total Single Revenue Vs Annualized For Top 20 Insurers L1(d)

def total_single_revenue_vs_annualized_for_top_20_insurers_L1d(df_dict_L1_d):
    # Loop through the loaded DataFrames and perform analysis for each file and sheet
    for (file, sheet_name), df in df_dict_L1_d.items():
        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Exclude the 'Market Total' row if it exists in the DataFrame
        df = df[df['Name of Insurer'] != 'Market Total']

        # Find the columns containing 'Channel Distribution: Total; Single Revenue Premiums (thousands of HKD)' and 'Channel Distribution: Total; Annualized Premiums (thousands of HKD)'
        total_single_revenue_column = df.columns[df.columns.str.contains('Channel Distribution:\s*Total\s*Single\s*Revenue\s*Premiums', case=False, na=False)].tolist()[0]
        total_annualized_premiums_column = df.columns[df.columns.str.contains('Channel Distribution:\s*Total\s*Annualized\s*Premiums', case=False, na=False)].tolist()[0]

        # Calculate the total single revenue premiums and total annualized premiums for each insurer and sort in descending order
        total_single_revenue_by_insurer = df.groupby('Name of Insurer')[total_single_revenue_column].sum().sort_values(ascending=False)
        total_annualized_premiums_by_insurer = df.groupby('Name of Insurer')[total_annualized_premiums_column].sum().sort_values(ascending=False)

        # Get the top 20 insurers
        top_20_insurers = total_single_revenue_by_insurer.head(20)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print(f"Top 20 Insurers - Channel Distribution: Total Single Revenue Premiums (thousands of HKD):")
        print(total_single_revenue_by_insurer.head(20))
        print(f"Top 20 Insurers - Channel Distribution: Total Annualized Premiums (thousands of HKD):")
        print(total_annualized_premiums_by_insurer.head(20))

        # Create a grouped bar chart to compare the Total Single Revenue Premiums and Total Annualized Premiums for the top 20 insurers
        plt.figure(figsize=(12, 6))
        ax = plt.gca()
        width = 0.35
        x = range(len(top_20_insurers))
        ax.bar(x, total_single_revenue_by_insurer.head(20), width, label='Total Single Revenue Premiums')
        ax.bar([i + width for i in x], total_annualized_premiums_by_insurer.head(20), width, label='Total Annualized Premiums')
        plt.xlabel('Insurer')
        plt.ylabel('Total Revenue (thousands of HKD)')
        plt.title(f'Top 20 Insurers - Channel Distribution: Total Single Revenue Premiums vs. Total Annualized Premiums for {file} - Sheet: {sheet_name}')
        plt.xticks([i + width/2 for i in x], top_20_insurers.index, rotation=90)
        plt.legend()

        # Format y-axis labels to display actual figures in thousands (K)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(format_thousands))

        plt.show()    
        
#Market Share Totals

def market_share_totals_L1d(df_dict_L1_d):
    for key, df in df_dict_L1_d.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on the provided formula and add it as a new column
        total_revenue_formula = (df['Channel Distribution: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Channel Distribution: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print("Total Revenue by Insurer:")
        print(total_revenue_by_insurer)
        print("\nMarket Share by Insurer (%):")
        print(market_share_by_insurer)
        print("\n")

       
        
# Distribution Channels Market Share Top 10 Analysis Pie Chart

def market_share_top_10_pie_chart_L1d(df_dict_L1_d):
    for key, df in df_dict_L1_d.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on the provided formula and add it as a new column
        total_revenue_formula = (df['Channel Distribution: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Channel Distribution: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Separate the top 10 insurers and combine the rest as "Others"
        top_10_insurers = market_share_by_insurer.head(10)
        others_market_share = market_share_by_insurer.iloc[10:].sum()
        top_10_insurers['Others'] = others_market_share

        # Plotting the pie chart
        plt.figure(figsize=(10, 6))
        plt.pie(top_10_insurers, labels=top_10_insurers.index, autopct='%1.1f%%', startangle=140, pctdistance=0.85)
        plt.axis('equal')
        plt.title(f"Distribution Channels Market Share Analysis - {file} - {sheet_name}")

        # Show the plot
        plt.show()

        
# Distribution Channels Market Share Top 10 Insurers Analysis Horizontal Bar Chart

def market_share_top_10_Horizontal_bar_L1d(df_dict_L1_d):
    for key, df in df_dict_L1_d.items():
        file, sheet_name = key

        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Calculate 'Total Revenue' based on the provided formula and add it as a new column
        total_revenue_formula = (df['Channel Distribution: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Channel Distribution: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()

        # Calculate the total market revenue
        total_market_revenue = total_revenue_by_insurer.sum()

        # Calculate the market share for each insurer
        market_share_by_insurer = (total_revenue_by_insurer / total_market_revenue) * 100

        # Sort the market share in descending order to get the dominant players first
        market_share_by_insurer = market_share_by_insurer.sort_values(ascending=False)

        # Separate the top 20 insurers and combine the rest as "Others"
        top_20_insurers = market_share_by_insurer.head(20)
        others_market_share = market_share_by_insurer.iloc[20:].sum()
        top_20_insurers['Others'] = others_market_share

        # Plotting the horizontal bar chart
        plt.figure(figsize=(10, 6))
        top_20_insurers.plot(kind='barh', color='skyblue', edgecolor='black')
        plt.xlabel('Market Share (%)')
        plt.ylabel('Insurer')
        plt.title(f"Onshore/Offshore Market Share Analysis - {file} - {sheet_name}")

        # Show the plot
        plt.show()

       
#Premiums Comparisons For Top 10 Insurers Stacked Bar Chart

def premium_comparison_top_10_insurers_L1d(df_dict_L1_d):
    for key, df in df_dict_L1_d.items():
        file, sheet_name = key

        # Filter out 'Market Total' rows
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate 'Total Revenue' based on the provided formula and add it as a new column
        total_revenue_formula = (df['Channel Distribution: Total Single Revenue Premiums (thousands of HKD)'] / 10) + df['Channel Distribution: Total Annualized Premiums (thousands of HKD)']
        df['Total Revenue'] = total_revenue_formula

        # Find the columns containing 'Premiums' and 'Revenues' (ignoring case and whitespace)
        annualized_columns = df.columns[df.columns.str.contains('Annualized', case=False, na=False)].tolist()
        single_columns = df.columns[df.columns.str.contains('Single', case=False, na=False)].tolist()

        # Drop the unwanted columns
        unwanted_columns = [
            'Channel Distribution: Total Single Revenue Premiums (thousands of HKD)',
            'Channel Distribution: Total Annualized Premiums (thousands of HKD)',
            'Total Revenue (thousands of HKD)']
        
        annualized_columns = [col for col in annualized_columns if col not in unwanted_columns]
        single_columns = [col for col in single_columns if col not in unwanted_columns]

        # Calculate the total revenue for each insurer and sort the insurers based on their total revenue
        total_revenue_by_insurer = df.groupby('Name of Insurer')['Total Revenue'].sum()
        sorted_top_ten_insurers = total_revenue_by_insurer.sort_values(ascending=False).head(10).index.tolist()

        # Select only the relevant columns for the Premium and Revenue Comparison and filter top ten insurers
        relevant_columns = ['Name of Insurer'] + annualized_columns + single_columns + ['Total Revenue']
        df = df[relevant_columns]
        df_top_ten = df[df['Name of Insurer'].isin(sorted_top_ten_insurers)].copy()

        # Drop the Total Revenue column from the DataFrame for plotting
        df_top_ten.drop(columns=['Total Revenue'], inplace=True)

        # Plotting the stacked bar chart for top ten insurers
        plt.figure(figsize=(12, 8))
        ax = df_top_ten.set_index('Name of Insurer')[annualized_columns + single_columns].plot(kind='bar', stacked=True)
        plt.xlabel('Insurer')
        plt.ylabel('Amount (thousands of HKD)')
        plt.title(f"Annualized and Single Comparison for Top Ten Insurers - {file} - {sheet_name}")

        # Format the y-axis ticks with thousands separator
        ax.get_yaxis().set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

        # Move the legend outside the plot area to avoid blocking
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1))

        # Show the plot
        plt.show()
        
#Total Revenue Comparison for Top 20 Insurers

def Total_Revenue_Comparison_For_Top_20_Insurers_L1d_Bar_Chart(df_dict_d):
    # Loop through the loaded DataFrames and perform analysis for each file and sheet
    for (file, sheet_name), df in df_dict_d.items():
        # Clean up column names by stripping whitespaces
        df.columns = df.columns.str.strip()

        # Exclude the 'Market Total' row if it exists in the DataFrame
        df = df[df['Name of Insurer'] != 'Market Total']

        # Calculate the total revenue for each insurer based on the formula
        df['Total Revenue'] = df['Channel Distribution: Total Single Revenue Premiums (thousands of HKD)'] / 10 + df['Channel Distribution: Total Annualized Premiums (thousands of HKD)']

        # Find the column containing 'Total Revenue' (ignoring case and whitespace)
        total_revenue_column = 'Total Revenue'

        # Calculate the total revenue for each insurer and sort in descending order
        total_revenue_by_insurer = df.groupby('Name of Insurer')[total_revenue_column].sum().sort_values(ascending=False)

        # Get the top 20 insurers
        top_20_insurers = total_revenue_by_insurer.head(20)

        # Print or store the results as per your requirement
        print(f"Analysis for {file} - Sheet: {sheet_name}:")
        print(f"Top 20 Insurers - Total Revenue (thousands of HKD):")
        print(top_20_insurers)

        # Create a bar chart to compare the Total Revenue for the top 20 insurers
        plt.figure(figsize=(12, 6))
        ax = top_20_insurers.plot(kind='bar')
        plt.xlabel('Insurer')
        plt.ylabel('Total Revenue (thousands of HKD)')
        plt.title(f'Top 20 Insurers - Total Revenue for {file} - Sheet: {sheet_name}')

        # Format y-axis labels to display actual figures in thousands (K)
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(format_thousands))

        plt.show()  
        
        
# User Interface
def user_interface(df_dict_L1, df_dict_L1_b, df_dict_L1_d):
    while True:
        print("\n===== Insurance Data Analysis =====")
        print("Select the Table to analyze:")
        print("1. Table L1")
        print("2. Table L1(b)")
        print("3. Table L1(d)")
        print("0. Exit")  # New choice to exit

        choice_table = input("Enter the option (0/1/2/3): ")

        if choice_table == '0':
            print("Exiting the program.")
            break  # Exit the loop and stop the program

        elif choice_table == '1':
            selected_df_dict = df_dict_L1
            table_name = "Table L1"
            suffix = "L1"

        elif choice_table == '2':
            selected_df_dict = df_dict_L1_b
            table_name = "Table L1(b)"
            suffix = "L1b"

        elif choice_table == '3':
            selected_df_dict = df_dict_L1_d
            table_name = "Table L1(d)"
            suffix = "L1d"

        else:
            print("Invalid option. Please try again.")
            continue

        print(f"\nSelected Table: {table_name}\n")
        
        while True:  # Inner loop for analysis selection
            print("Select the type of analysis:")
            
            # Mapping available analysis options to the respective DataFrame dictionaries
            available_analysis_options = {
                '1': [fn for fn in [Total_Revenue_Comparison_For_Top_20_Insurers_L1,
                                    Single_Vs_Annualized_Comparison_Bar_Chart_L1,
                                    Market_Share_Totals_L1,
                                    Market_Share_Top_10_Pie_Chart_L1,
                                    Market_Share_Top_10_Horizontal_Bar_Chart_L1,
                                    Classes_Comparisons_L1_Bar_Chart] if fn.__name__.endswith(suffix)],
                '2': [fn for fn in [Total_Revenue_Comparison_For_Top_20_Insurers_L1b,
                                    Onshore_Offshore_Market_Share_Top_20_Insurers_Pie_Chart_L1b, 
                                    Market_Share_Totals_L1b,
                                    Onshore_Offshore_Market_Share_Top_20_Insurers_Horizontal_Bar_L1b,
                                    Premium_Comparison_For_Top_10_Insurers_L1b] if fn.__name__.endswith(suffix)],
                '3': [fn for fn in [market_share_totals_L1d, 
                                    market_share_top_10_Horizontal_bar_L1d, 
                                    Total_Revenue_Comparison_For_Top_20_Insurers_L1d_Bar_Chart, 
                                    total_single_revenue_vs_annualized_for_top_20_insurers_L1d,
                                    market_share_top_10_pie_chart_L1d,
                                    premium_comparison_top_10_insurers_L1d] if fn.__name__.endswith(suffix)],
            }

            # Display available analysis options for the selected Table
            for index, analysis_func in enumerate(available_analysis_options[choice_table], start=1):
                print(f"{index}. {analysis_func.__name__.replace('_', ' ').title()}")

            # Exit option added to the inner loop
            print("0. Exit to Table selection")

            choice_analysis = input("Enter the option (0 or the corresponding number): ")

            if choice_analysis == '0':
                break  # Exit the inner loop and go back to table selection

            elif choice_analysis.isdigit() and int(choice_analysis) in range(1, len(available_analysis_options[choice_table]) + 1):
                analysis_func = available_analysis_options[choice_table][int(choice_analysis) - 1]
                analysis_func(selected_df_dict)
            
            else:
                print("Invalid option. Please try again.")
        
        print("\nDo you want to perform another analysis for the same table?")
        print("1. Yes")
        print("2. No")
        repeat_choice = input("Enter the option (1 or 2): ")

        if repeat_choice == '1':
            continue  # Repeat analysis for the same table

        elif repeat_choice == '2':
            print("Going back to Table selection.")
            continue  # Go back to table selection

        else:
            print("Invalid option. Going back to Table selection.")
            continue

if __name__ == "__main__":
    # Prompt user to input start and end years and quarters
    start_year = int(input("Enter the start year (e.g., 2014): "))
    start_quarter = int(input("Enter the start quarter (1 to 4): "))
    end_year = int(input("Enter the end year (e.g., 2016): "))
    end_quarter = int(input("Enter the end quarter (1 to 4): "))

    # Load data for Table L1, Table L1(b), and Table L1(d)
    data_directory = 'C:\\Users\\User\\Desktop\\Data Insurance\\output'
    df_dict_L1, df_dict_L1_b, df_dict_L1_d = load_insurance_data(data_directory, start_year, start_quarter, end_year, end_quarter)
    print("Data loaded successfully.")

    # Call the user interface function to start the analysis
    user_interface(df_dict_L1, df_dict_L1_b, df_dict_L1_d)




```

    Enter the start year (e.g., 2014):  2018
    Enter the start quarter (1 to 4):  2
    Enter the end year (e.g., 2016):  2028
    Enter the end quarter (1 to 4):  4
    

    Data loaded successfully.
    
    ===== Insurance Data Analysis =====
    Select the Table to analyze:
    1. Table L1
    2. Table L1(b)
    3. Table L1(d)
    0. Exit
    

    Enter the option (0/1/2/3):  1
    

    
    Selected Table: Table L1
    
    Select the type of analysis:
    1. Total Revenue Comparison For Top 20 Insurers L1
    2. Single Vs Annualized Comparison Bar Chart L1
    3. Market Share Totals L1
    4. Market Share Top 10 Pie Chart L1
    5. Market Share Top 10 Horizontal Bar Chart L1
    0. Exit to Table selection
    

    Enter the option (0 or the corresponding number):  5
    


    
![png](output_3_5.png)
    



    
![png](output_3_6.png)
    



    
![png](output_3_7.png)
    



    
![png](output_3_8.png)
    



    
![png](output_3_9.png)
    



    
![png](output_3_10.png)
    



    
![png](output_3_11.png)
    



    
![png](output_3_12.png)
    



    
![png](output_3_13.png)
    



    
![png](output_3_14.png)
    



    
![png](output_3_15.png)
    



    
![png](output_3_16.png)
    



    
![png](output_3_17.png)
    



    
![png](output_3_18.png)
    



    
![png](output_3_19.png)
    



    
![png](output_3_20.png)
    



    
![png](output_3_21.png)
    



    
![png](output_3_22.png)
    



    
![png](output_3_23.png)
    



    
![png](output_3_24.png)
    


    Select the type of analysis:
    1. Total Revenue Comparison For Top 20 Insurers L1
    2. Single Vs Annualized Comparison Bar Chart L1
    3. Market Share Totals L1
    4. Market Share Top 10 Pie Chart L1
    5. Market Share Top 10 Horizontal Bar Chart L1
    0. Exit to Table selection
    

    Enter the option (0 or the corresponding number):  0
    

    
    Do you want to perform another analysis for the same table?
    1. Yes
    2. No
    

    Enter the option (1 or 2):  2
    

    Going back to Table selection.
    
    ===== Insurance Data Analysis =====
    Select the Table to analyze:
    1. Table L1
    2. Table L1(b)
    3. Table L1(d)
    0. Exit
    

    Enter the option (0/1/2/3):  1
    

    
    Selected Table: Table L1
    
    Select the type of analysis:
    1. Total Revenue Comparison For Top 20 Insurers L1
    2. Single Vs Annualized Comparison Bar Chart L1
    3. Market Share Totals L1
    4. Market Share Top 10 Pie Chart L1
    5. Market Share Top 10 Horizontal Bar Chart L1
    0. Exit to Table selection
    

    Enter the option (0 or the corresponding number):  2
    

    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2018 Q2.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                    Name of Insurer  Total Revenue
    31                    HSBC Life      9182803.4
    1             AIA International      9039366.1
    14                   China Life      8410713.0
    45         Prudential (HK) Life      7997263.2
    12                     BOC LIFE      6223399.0
    26          Hang Seng Insurance      2874213.8
    15                        TPLHK      2824904.0
    33             Manulife (Int'l)      2155426.2
    6           AXA China (Bermuda)      1540765.0
    22                     FWD Life      1246061.3
    20                       FTLife       914975.7
    10                     BEA Life       771233.2
    56                         TLIC       708846.9
    55           Sun Life Hong Kong       578788.7
    35              MassMutual Asia       422088.4
    21         Fubon Life Hong Kong       374614.0
    36                      MetLife       284324.2
    29               Hong Kong Life       240403.1
    16                   Chubb Life       159282.9
    58  Transamerica Life (Bermuda)       157777.0
    


    
![png](output_3_35.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2018 Q3.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                    Name of Insurer  Total Revenue
    1             AIA International     13696695.1
    31                    HSBC Life     13455877.8
    45         Prudential (HK) Life     12263446.8
    14                   China Life     11513562.0
    12                     BOC LIFE      7839824.2
    26          Hang Seng Insurance      4025931.8
    15                        TPLHK      3573322.0
    33             Manulife (Int'l)      3411361.3
    6           AXA China (Bermuda)      2410491.8
    22                     FWD Life      1823181.2
    20                       FTLife      1351746.5
    10                     BEA Life      1250855.5
    56                         TLIC      1113220.4
    55           Sun Life Hong Kong       826461.9
    35              MassMutual Asia       623796.8
    21         Fubon Life Hong Kong       599042.0
    36                      MetLife       442193.1
    29               Hong Kong Life       356952.7
    16                   Chubb Life       291514.8
    58  Transamerica Life (Bermuda)       256899.6
    


    
![png](output_3_38.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2018 Q4.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                    Name of Insurer  Total Revenue
    1             AIA International     19334144.7
    45         Prudential (HK) Life     17759641.8
    32                    HSBC Life     16563057.5
    15                   China Life     14281375.0
    12                     BOC LIFE      9240390.3
    34             Manulife (Int'l)      4784119.8
    27          Hang Seng Insurance      4772898.9
    16                        TPLHK      3908463.0
    6           AXA China (Bermuda)      3503741.1
    23                     FWD Life      2822586.1
    21                       FTLife      1942726.9
    10                     BEA Life      1614444.8
    56                         TLIC      1563615.5
    55           Sun Life Hong Kong      1261542.4
    59                      YF LIFE       887452.2
    22         Fubon Life Hong Kong       832444.0
    36                      MetLife       644683.0
    17                   Chubb Life       468393.1
    30               Hong Kong Life       441708.2
    58  Transamerica Life (Bermuda)       293230.1
    


    
![png](output_3_41.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2019 Q1.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life      9203957.0
    31             HSBC Life      5406111.0
    1      AIA International      5191984.1
    44  Prudential (HK) Life      4150278.8
    12              BOC LIFE      2525074.2
    16                 TPLHK      2022695.0
    26   Hang Seng Insurance      1958870.9
    33      Manulife (Int'l)      1236771.5
    23              FWD Life      1044394.2
    6    AXA China (Bermuda)      1011011.7
    10              BEA Life       993365.8
    55                  TLIC       504700.4
    21                FTLife       434483.1
    22  Fubon Life Hong Kong       388798.0
    29        Hong Kong Life       292500.2
    54    Sun Life Hong Kong       272952.1
    59               YF LIFE       186185.2
    17            Chubb Life       152205.1
    35               MetLife       136017.2
    25    Generali Life (HK)        65836.8
    


    
![png](output_3_44.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2019 Q2.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     19174101.0
    31             HSBC Life     10122958.5
    1      AIA International      9956617.9
    44  Prudential (HK) Life      8426897.4
    12              BOC LIFE      7296037.3
    16                 TPLHK      4041968.0
    26   Hang Seng Insurance      3073456.5
    33      Manulife (Int'l)      2721975.3
    6    AXA China (Bermuda)      2111594.2
    23              FWD Life      1845324.6
    10              BEA Life      1835009.8
    56                  TLIC      1043698.4
    22  Fubon Life Hong Kong      1020786.0
    21                FTLife       922338.2
    54    Sun Life Hong Kong       712310.0
    29        Hong Kong Life       416692.7
    61               YF LIFE       388906.4
    17            Chubb Life       295994.1
    35               MetLife       258499.4
    25    Generali Life (HK)       141432.5
    


    
![png](output_3_47.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2019 Q3.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     22340664.0
    31             HSBC Life     14890357.6
    1      AIA International     13313447.9
    12              BOC LIFE     12163619.4
    45  Prudential (HK) Life     11742134.0
    16                 TPLHK      5529210.0
    33      Manulife (Int'l)      4724817.3
    26   Hang Seng Insurance      3911753.2
    6    AXA China (Bermuda)      2992772.7
    23              FWD Life      2473607.3
    10              BEA Life      2285631.8
    22  Fubon Life Hong Kong      1517093.0
    21                FTLife      1377265.6
    57                  TLIC      1358793.5
    55    Sun Life Hong Kong      1243017.2
    62               YF LIFE      1134649.2
    29        Hong Kong Life       495339.4
    17            Chubb Life       391432.1
    35               MetLife       377292.0
    25    Generali Life (HK)       179702.2
    


    
![png](output_3_50.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2019 Q4.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     24548158.0
    32             HSBC Life     17715256.4
    1      AIA International     17192582.3
    46  Prudential (HK) Life     15789332.6
    12              BOC LIFE     13805808.2
    34      Manulife (Int'l)      6413488.4
    17                 TPLHK      5998533.0
    27   Hang Seng Insurance      4456100.8
    6    AXA China (Bermuda)      3737922.2
    24              FWD Life      3292529.6
    10              BEA Life      2431880.5
    22                FTLife      1996174.3
    56    Sun Life Hong Kong      1946728.3
    58                  TLIC      1730984.8
    23  Fubon Life Hong Kong      1543321.0
    63               YF LIFE      1420967.8
    30        Hong Kong Life       556033.4
    18            Chubb Life       511395.8
    36               MetLife       488061.0
    26    Generali Life (HK)       195964.6
    


    
![png](output_3_53.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2020 Q1.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life      3657090.0
    32             HSBC Life      3532731.7
    12              BOC LIFE      3356033.1
    1      AIA International      2451580.0
    17                 TPLHK      2200981.0
    45  Prudential (HK) Life      2055167.2
    34      Manulife (Int'l)      1486258.8
    27   Hang Seng Insurance      1041712.8
    5    AXA China (Bermuda)       608849.4
    24              FWD Life       530655.6
    56    Sun Life Hong Kong       491441.0
    9               BEA Life       488693.8
    22                FTLife       422925.5
    58                  TLIC       378338.8
    63               YF LIFE       159658.8
    30        Hong Kong Life       137268.7
    26    Generali Life (HK)        77350.1
    18            Chubb Life        76922.9
    23  Fubon Life Hong Kong        36921.0
    29          HKMC Annuity        32017.4
    


    
![png](output_3_56.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2020 Q2.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life      8563148.0
    32             HSBC Life      6489196.3
    12              BOC LIFE      5442281.7
    1      AIA International      3814475.4
    17                 TPLHK      3304403.0
    45  Prudential (HK) Life      3009129.4
    34      Manulife (Int'l)      2830026.0
    27   Hang Seng Insurance      1702495.3
    9               BEA Life      1425327.8
    24              FWD Life      1268883.6
    56    Sun Life Hong Kong       966173.4
    5    AXA China (Bermuda)       957202.1
    22                FTLife       807330.4
    58                  TLIC       639523.7
    63               YF LIFE       394754.6
    30        Hong Kong Life       268796.9
    26    Generali Life (HK)       188924.4
    18            Chubb Life       143248.3
    23  Fubon Life Hong Kong       130523.0
    29          HKMC Annuity        92196.0
    


    
![png](output_3_59.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2020 Q3.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     12782595.0
    35             HSBC Life      9563025.3
    12              BOC LIFE      8996235.0
    1      AIA International      5466382.5
    37      Manulife (Int'l)      4276691.0
    17                 TPLHK      4157092.0
    46  Prudential (HK) Life      4106365.0
    29   Hang Seng Insurance      2517410.9
    9               BEA Life      2164489.1
    25    FWD Life (Bermuda)      1903416.4
    56    Sun Life Hong Kong      1548108.1
    5    AXA China (Bermuda)      1384457.7
    22                FTLife      1210765.8
    58                  TLIC       768258.8
    63               YF LIFE       678493.3
    33        Hong Kong Life       369000.5
    23  Fubon Life Hong Kong       231942.0
    28    Generali Life (HK)       206988.4
    18            Chubb Life       205663.5
    32          HKMC Annuity       145846.7
    


    
![png](output_3_62.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2020 Q4.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     13895216.0
    35             HSBC Life     11599291.9
    12              BOC LIFE     11344367.9
    1      AIA International      7254380.0
    46  Prudential (HK) Life      5880072.4
    37      Manulife (Int'l)      5655059.1
    17                 TPLHK      5213810.1
    29   Hang Seng Insurance      3092272.9
    25    FWD Life (Bermuda)      2727803.9
    55    Sun Life Hong Kong      2457628.7
    9               BEA Life      2426533.8
    5    AXA China (Bermuda)      1837247.5
    22                FTLife      1739192.9
    62               YF LIFE       884800.9
    57                  TLIC       805352.5
    33        Hong Kong Life       435894.6
    23  Fubon Life Hong Kong       324887.0
    18            Chubb Life       285613.7
    32          HKMC Annuity       253792.8
    28    Generali Life (HK)       237910.1
    


    
![png](output_3_65.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2021 Q1.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life      4436818.0
    12              BOC LIFE      3968454.7
    35             HSBC Life      3544331.5
    1      AIA International      1605047.6
    37      Manulife (Int'l)      1523699.7
    17                 TPLHK      1378856.0
    9               BEA Life      1209166.0
    29   Hang Seng Insurance      1104029.7
    46  Prudential (HK) Life      1076151.6
    25    FWD Life (Bermuda)       785854.6
    5    AXA China (Bermuda)       577003.5
    55    Sun Life Hong Kong       488219.9
    22                FTLife       482134.7
    33        Hong Kong Life       257434.3
    23  Fubon Life Hong Kong       243821.0
    62               YF LIFE       175161.5
    32          HKMC Annuity        88796.4
    18            Chubb Life        66190.5
    61        Well Link Life        49647.0
    28    Generali Life (HK)        37191.0
    


    
![png](output_3_68.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2021 Q2.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life      8652773.0
    35             HSBC Life      7652324.9
    12              BOC LIFE      6286261.8
    1      AIA International      2954157.2
    37      Manulife (Int'l)      2879826.9
    17                 TPLHK      2786940.0
    46  Prudential (HK) Life      1962577.5
    29   Hang Seng Insurance      1873886.5
    25    FWD Life (Bermuda)      1634174.1
    9               BEA Life      1247440.1
    5    AXA China (Bermuda)      1106631.6
    22                FTLife      1024431.8
    55    Sun Life Hong Kong       849012.5
    23  Fubon Life Hong Kong       536713.0
    33        Hong Kong Life       391358.7
    62               YF LIFE       361536.1
    32          HKMC Annuity       157514.5
    18            Chubb Life       127793.2
    61        Well Link Life        88336.7
    28    Generali Life (HK)        64194.3
    


    
![png](output_3_71.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2021 Q3.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     13029252.0
    35             HSBC Life     11186991.1
    12              BOC LIFE      8320140.7
    2      AIA International      4493622.6
    37      Manulife (Int'l)      4416217.2
    17                 TPLHK      4123475.0
    46  Prudential (HK) Life      3008900.3
    29   Hang Seng Insurance      2698378.8
    25    FWD Life (Bermuda)      2411492.0
    6    AXA China (Bermuda)      1585602.7
    22                FTLife      1432056.3
    1            AIA Everest      1249322.1
    55    Sun Life Hong Kong      1113276.3
    23  Fubon Life Hong Kong      1064462.0
    62               YF LIFE       653618.6
    33        Hong Kong Life       464328.8
    32          HKMC Annuity       217512.1
    18            Chubb Life       191626.6
    28    Generali Life (HK)        96883.8
    61        Well Link Life        95320.1
    


    
![png](output_3_74.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2021 Q4.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life     14976924.0
    36             HSBC Life     13926100.1
    12              BOC LIFE     10283521.2
    2      AIA International      6573626.7
    38      Manulife (Int'l)      5929749.9
    18                 TPLHK      5860483.0
    47  Prudential (HK) Life      4274877.3
    30   Hang Seng Insurance      3210753.2
    26    FWD Life (Bermuda)      3170722.2
    6    AXA China (Bermuda)      2313829.0
    23                FTLife      2074978.7
    56    Sun Life Hong Kong      1592029.5
    24  Fubon Life Hong Kong      1316382.0
    1            AIA Everest      1250048.0
    63               YF LIFE       979247.8
    34        Hong Kong Life       568686.8
    33          HKMC Annuity       300344.3
    19            Chubb Life       269479.4
    7         AXA China (HK)       258488.0
    29    Generali Life (HK)       138377.5
    


    
![png](output_3_77.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2022 Q1.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                          Name of Insurer  Total Revenue
    12                           BOC LIFE      3629517.3
    36                          HSBC Life      3162136.6
    15                         China Life      2485670.0
    38                   Manulife (Int'l)      1163575.1
    18                              TPLHK      1160121.0
    2                   AIA International      1123293.0
    47               Prudential (HK) Life       836109.7
    30                Hang Seng Insurance       751837.0
    24               Fubon Life Hong Kong       660231.0
    26                 FWD Life (Bermuda)       596482.7
    6                 AXA China (Bermuda)       367612.3
    23                             FTLife       268890.2
    56                 Sun Life Hong Kong       260716.0
    63                            YF LIFE       143826.5
    29                 Generali Life (HK)        67075.0
    34                     Hong Kong Life        65835.9
    19                         Chubb Life        55555.4
    33                       HKMC Annuity        43787.5
    67  Zurich Life Insurance (Hong Kong)        29109.0
    64                          ZA Insure        26350.8
    


    
![png](output_3_80.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2022 Q2.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                          Name of Insurer  Total Revenue
    12                           BOC LIFE      6603985.0
    36                          HSBC Life      6022545.4
    15                         China Life      3861321.0
    2                   AIA International      2482643.7
    38                   Manulife (Int'l)      2115982.0
    18                              TPLHK      1890605.0
    47               Prudential (HK) Life      1644472.6
    30                Hang Seng Insurance      1294600.9
    24               Fubon Life Hong Kong      1161935.0
    26                 FWD Life (Bermuda)       943300.1
    6                 AXA China (Bermuda)       737814.5
    23                             FTLife       701628.6
    56                 Sun Life Hong Kong       515151.1
    63                            YF LIFE       352360.4
    29                 Generali Life (HK)       153125.0
    19                         Chubb Life       139039.8
    33                       HKMC Annuity       121350.5
    34                     Hong Kong Life       112684.6
    67  Zurich Life Insurance (Hong Kong)        50037.0
    62                     Well Link Life        36817.8
    


    
![png](output_3_83.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2022 Q3.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                          Name of Insurer  Total Revenue
    36                          HSBC Life      8712633.5
    12                           BOC LIFE      7616928.9
    15                         China Life      5046745.0
    2                   AIA International      3905840.0
    38                   Manulife (Int'l)      3058786.8
    47               Prudential (HK) Life      2626998.7
    18                              TPLHK      2148317.0
    30                Hang Seng Insurance      1803555.8
    24               Fubon Life Hong Kong      1406195.0
    26                 FWD Life (Bermuda)      1250640.2
    6                 AXA China (Bermuda)      1159405.0
    23                             FTLife      1092134.2
    56                 Sun Life Hong Kong       847689.7
    63                            YF LIFE       629999.7
    29                 Generali Life (HK)       250581.0
    19                         Chubb Life       233073.8
    33                       HKMC Annuity       196857.2
    34                     Hong Kong Life       160371.0
    67  Zurich Life Insurance (Hong Kong)        74750.0
    62                     Well Link Life        44734.9
    


    
![png](output_3_86.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2022 Q4.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
                          Name of Insurer  Total Revenue
    36                          HSBC Life     10089565.1
    12                           BOC LIFE      8713030.0
    2                   AIA International      5877057.5
    15                         China Life      5780805.7
    38                   Manulife (Int'l)      3914999.9
    47               Prudential (HK) Life      3884031.3
    18                              TPLHK      2308575.0
    30                Hang Seng Insurance      2107523.4
    26                 FWD Life (Bermuda)      1619956.1
    6                 AXA China (Bermuda)      1597883.5
    24               Fubon Life Hong Kong      1436641.0
    23                             FTLife      1428531.9
    55                 Sun Life Hong Kong      1246972.4
    63                            YF LIFE       897276.1
    19                         Chubb Life       365765.1
    29                 Generali Life (HK)       356826.0
    33                       HKMC Annuity       252873.3
    34                     Hong Kong Life       200567.4
    67  Zurich Life Insurance (Hong Kong)       105788.0
    13                        Bowtie Life        59557.0
    


    
![png](output_3_89.png)
    


    C:\Users\User\AppData\Local\Temp\ipykernel_12788\538755058.py:130: SettingWithCopyWarning: 
    A value is trying to be set on a copy of a slice from a DataFrame.
    Try using .loc[row_indexer,col_indexer] = value instead
    
    See the caveats in the documentation: https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
      df['Total Revenue'] = df[total_single_revenue_column] / 10 + df[total_annualized_premiums_column]
    

    Analysis for 2023 Q1.xlsx - Sheet: Table L1:
    Top 20 Insurers - Linked/Non-Linked: Total Revenue (thousands of HKD):
             Name of Insurer  Total Revenue
    15            China Life      5935949.8
    36             HSBC Life      3567227.6
    2      AIA International      3505173.5
    47  Prudential (HK) Life      3494090.2
    12              BOC LIFE      2747041.7
    38      Manulife (Int'l)      1422710.0
    30   Hang Seng Insurance      1413142.9
    6    AXA China (Bermuda)       930575.1
    26    FWD Life (Bermuda)       726788.0
    23                FTLife       594774.9
    55    Sun Life Hong Kong       391464.4
    63               YF LIFE       388767.1
    18                 TPLHK       328909.0
    29    Generali Life (HK)       122458.0
    19            Chubb Life       108944.1
    24  Fubon Life Hong Kong       104915.0
    34        Hong Kong Life        68104.4
    11                  Blue        59848.3
    62        Well Link Life        36617.7
    33          HKMC Annuity        33912.5
    


    
![png](output_3_92.png)
    


    Select the type of analysis:
    1. Total Revenue Comparison For Top 20 Insurers L1
    2. Single Vs Annualized Comparison Bar Chart L1
    3. Market Share Totals L1
    4. Market Share Top 10 Pie Chart L1
    5. Market Share Top 10 Horizontal Bar Chart L1
    0. Exit to Table selection
    


```python

```


```python

```
