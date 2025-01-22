import polars as pl

def read_excel_file(file):
    return pl.read_excel(file)