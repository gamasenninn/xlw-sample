import xlwings as xw
import pandas as pd
import yfinance as yf

def get_market_data():
    sheet = wb.sheets["src"]
    sheet_conf = wb.sheets["conf"]
    code = sheet_conf["B1"].value
    if sheet_conf["B2"].value == 1:
        ticker_info = yf.Ticker(code)
        hist = ticker_info.history(period="max")
        sheet.clear()
        sheet.range('A1').options(pd.DataFrame, header=1,index=True).value = hist

def simulation():
    sheet_src = wb.sheets["src"]
    sheet_conf = wb.sheets["conf"]
    df_conf = sheet_conf.range('D2').expand('table').options(pd.DataFrame, header=1,index=False).value
    for idx,row in df_conf.iterrows():
        sid ,start, end  = [row['ID'],row['開始期間'],row['終了期間']]
        sheet_dst = wb.sheets[sid] if sid in [sh.name for sh in wb.sheets] else wb.sheets.add(sid)
        sheet_dst.clear()
        df = sheet_src.range('A1').expand('table').options(pd.DataFrame, header=1,index=True).value
        df = df[(df.index>=start) & (df.index <= end)]
        sheet_dst.range('A1').options(pd.DataFrame, header=1,index=True).value = df
       
def main():
    global wb
    wb = xw.Book.caller()
    get_market_data()
    simulation()
    
if __name__ == "__main__":
    xw.Book("kabu.xlsx").set_mock_caller()
    main()
