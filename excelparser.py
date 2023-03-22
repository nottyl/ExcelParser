import pandas as pd


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

    # Write the modified data to a new Excel file
    with pd.ExcelWriter('parsed_file.xlsx') as writer:
        workbook.to_excel(writer, index=False)
        print("\n[Output generated!]\n")
        print(workbook.head())


def generate_new_file(input_file):
    new_data = {"六字學校": [], "報名賽制": [], "分隊": [], "身份": [], "中文名字": [], "English Name": []}

    df = pd.read_excel(input_file)

    for index, row in df.iterrows():
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
    for index, row in df.iterrows():

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
    for index, row in df.iterrows():
        for col in [7, 9, 11, 13, 15, 17, 3, 5]:
            col_value = row[col]
            if pd.isna(col_value):
                continue
            else:
                new_data["中文名字"][j] = col_value
                j += 1

    new_df = pd.DataFrame.from_dict(new_data)
    new_df.to_excel('output_file.xlsx', index=False)