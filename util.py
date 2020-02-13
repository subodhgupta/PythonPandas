import pandas as pd


def retrieve_column_connection(path):
    """
    get data from column connection spreadsheet.
    :param path: Column connections file path
    :return: Data frame with column connection data
    """
    return pd.read_excel(path)


def distinct_workbook_worksheet(df_column_connections):
    """
    Store distinct workbook/worksheet into a dataframe
    :param df_column_connections:
    :return: Dataframe
    """
    return df_column_connections[['workbook', 'worksheet']].groupby(['workbook', 'worksheet']).first().reset_index()
