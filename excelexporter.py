from datetime import datetime, timedelta
import pandas as pd

def export_dict_to_excel(path,dict):
    """
    Useful for testing dicts to excel.
    """
    all_dates = set()
    for stock_data in dict.values():
        all_dates.update(stock_data.keys())

    all_dates = sorted(all_dates, key=lambda date: datetime.strptime(date, '%m/%d/%Y'))

    df = pd.DataFrame(index=all_dates)

    # Fill the DataFrame with stock data
    for stock, stock_data in dict.items():
        df[stock] = df.index.map(stock_data).fillna(0.0)

    df = df.T

    # Export to Excel
    df.to_excel(path, engine='openpyxl')