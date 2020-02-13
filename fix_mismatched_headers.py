# Read excel files with mismatched column names and even incorrectly mapped columns
# For example, a column can be called "Student Name" in one file and "Student NM" in another file
# Also, "Student Name" can be the first column in one file but the third column in second file
# This program can read a mapping file and consolidate the mismatched data.
# import dependencies
from pandas import DataFrame, read_excel, ExcelWriter
import os
import util

# Get the folder name to build filepath
folder = "static"
consolidated_workbook = "all_students"

column_connections = util.retrieve_column_connection(os.path.join(folder, "mapping_file.xlsx"))
worksheets = util.distinct_workbook_worksheet(column_connections)

print(worksheets)

# Create dataframe to store final result
df_all_students = DataFrame()

# Iterate over dataframe using the itertuples method (Use Pandas documentation as reference)
for workbook, worksheet in worksheets.itertuples(index=False):
    # Filter workbook/worksheet and store the result into a dataframe
    columns_detail = column_connections[(column_connections["workbook"] == workbook) &
                                        (column_connections["worksheet"] == worksheet)]

    # Create column id list - we will use as a parameter to read_excel method below
    columns_id_list = columns_detail["column_id"].to_list()

    # Do the same for new names - column rename_to of column_connection
    columns_rename_list = columns_detail["rename_to"].to_list()

    # Console status - using recursive string method
    print(f"Start to read {workbook} -> {worksheet}")
    print(f"Columns ID List: {columns_id_list}")
    print(f"Rename Columns List: {columns_rename_list}")

    # Build file path using os.path.join method
    filepath = os.path.join(folder, workbook + ".xlsx")

    # sort columns based on indices
    columns_rename_list.sort(key=dict(zip(columns_rename_list, columns_id_list)).get)

    # Fill partial result using read_excel to read this specific workbook
    partial_result = read_excel(filepath,
                                sheet_name=worksheet,
                                usecols=columns_id_list,
                                names=columns_rename_list)

    # show the partial results

    df_all_students = df_all_students.append(partial_result, ignore_index=True)
    # print(df_all_students)

consolidated_workbook_filepath = os.path.join(folder, consolidated_workbook + ".xlsx")
writer = ExcelWriter(consolidated_workbook_filepath)
df_all_students.to_excel(writer, consolidated_workbook, index=False)
writer.save()

# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.itertuples.html#pandas.DataFrame.itertuples
