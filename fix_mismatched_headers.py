# import dependencies
import pandas as pd
import os
import util as u
from importlib import reload as rl

# Get the folder name to build filepath
folder = 'static'

column_connections = u.retrieve_column_connection(os.path.join('static', 'column_connections.xlsx'))
worksheets = u.distinct_workbook_worksheet(column_connections)

# Create dataframe to store final result
df_all_students = pd.DataFrame()

# Iterate over dataframe using the itertuples method (Use Pandas documentation as reference)
for workbook, worksheet in worksheets.itertuples(index=False):

    # Filter workbook/worksheet and store the result into a dataframe
    columns_detail = column_connections[(column_connections['workbook'] == workbook) &
                       (column_connections['worksheet'] == worksheet)]

    # Create column id list - we will use as a parameter to read_excel method below
    columns_id_list = columns_detail['column_id'].to_list()

    # Do the same for new names - column rename_to of column_connection
    columns_rename_list = columns_detail['rename_to'].to_list()

    # Console status - using recursive string method
    print('Start to read {workbook} -> {worksheet}'.format(workbook=workbook, worksheet=worksheet))
    print('Columns ID List:')
    print(columns_id_list)
    print('Rename Columns List:')
    print(columns_rename_list)

    # Build file path using os.path.join method
    filepath = os.path.join(folder, workbook + '.xlsx')

    # Fill partial result using read_excel to read this specific workbook
    partial_result = pd.read_excel(filepath,
                                   sheet_name=worksheet,
                                   usecols=columns_id_list,
                                   names=columns_rename_list)

    # show the partial results
    print(partial_result)

    df_all_students = df_all_students.append(partial_result, ignore_index=True)


# refências:
# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.itertuples.html#pandas.DataFrame.itertuples
