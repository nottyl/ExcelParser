import pandas as pd

# Initialization of the input_file for better parsing experience
def init(input_file):
    # Please rename the Excel file to read as 'input_file'
    workbook = pd.read_excel(input_file)
    workbook.head()
    print("[File loaded]")

    # Removing irrelevant columns for better parsing experience
    columns_to_keep = [2, 3, 5, 25, 26, 28, 29, 31, 32, 34, 35, 37, 38, 40, 41, 43, 44, 46, 47]
    print("[Deleting irrelevant columns]\n")
    workbook = workbook.iloc[:, columns_to_keep]
    print("Column names in workbook:", workbook.columns)

    with pd.ExcelWriter('parsed_file.xlsx') as writer:
        workbook.to_excel(writer, index=False)
        print("\n[Output generated!]\n")
        print(workbook.head())


# Main parsing function. Please always execute this function.
def filter_and_sort(input_file):
    print("[Filtering and sorting...]")
    new_data = {"六字學校": [], "報名賽制": [], "分隊": [], "身份": [], "中文名字": [], "English Name": []}

    workbook = pd.read_excel(input_file)

    for index, row in workbook.iterrows():
        num_columns = int((len(row.dropna()) - 3) / 2)

        col_2_value = row[1]
        col_3_value = row[2]

        new_values = [col_2_value] * num_columns

        for i in range(len(new_values)):
            new_data["六字學校"].append(col_2_value)
            new_data["報名賽制"].append(col_3_value)
            new_data["分隊"].append("")
            new_data["身份"].append("")
            new_data["中文名字"].append("")
            new_data["English Name"].append("")

    i = 0
    for index, row in workbook.iterrows():

        for col in [8, 10, 12, 14, 16, 18, 4, 6]:
            col_value = row[col]

            if pd.isna(col_value):
                continue
            else:
                if col == 8:
                    new_data["分隊"][i] = "小隊1"
                    new_data["身份"][i] = "正1"
                if col == 10:
                    new_data["分隊"][i] = "小隊1"
                    new_data["身份"][i] = "正2"
                if col == 12:
                    new_data["分隊"][i] = "小隊2"
                    new_data["身份"][i] = "正1"
                if col == 14:
                    new_data["分隊"][i] = "小隊2"
                    new_data["身份"][i] = "正2"
                if col == 16:
                    new_data["身份"][i] = "備1"
                if col == 18:
                    new_data["身份"][i] = "備2"
                if col == 4 or col == 6:
                    new_data["身份"][i] = "指導教師"
                new_data["English Name"][i] = col_value
                i += 1

    j = 0
    for index, row in workbook.iterrows():
        for col in [7, 9, 11, 13, 15, 17, 3, 5]:
            col_value = row[col]
            if pd.isna(col_value):
                continue
            else:
                new_data["中文名字"][j] = col_value
                j += 1

    new_workbook = pd.DataFrame.from_dict(new_data)
    new_workbook.to_excel('output_file.xlsx', index=False)
    print("[Filtered and sorted through the .xlsx file!]")


# Sort by categories then put individual categories into a new sheet
def split_categories(input_file):
    workbook = pd.read_excel(input_file)
    categories = workbook['報名賽制'].unique()
    with pd.ExcelWriter('output_file1.xlsx') as writer:
        for category in categories:
            category_sheet = pd.DataFrame(columns=workbook.columns)
            category_data = workbook.loc[workbook['報名賽制'] == category]
            category_sheet = pd.concat([category_sheet, category_data])
            category_sheet.to_excel(writer, sheet_name=category, index=False)
    print("[Splitting the workbook based on the categories]")


# Sort by categories in the same sheet
def sort_categories(input_file):
    workbook = pd.read_excel(input_file)
    workbook_sorted = workbook.sort_values(by=['報名賽制'], kind='mergesort')
    workbook_sorted.to_excel('output_file.xlsx', index=False)
    print("[Sorting the workbook based on the categories]")

#TODO: formatting
