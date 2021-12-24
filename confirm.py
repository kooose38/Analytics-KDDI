import pandas as pd 
import numpy as np 
import time 
import os 
import PySimpleGUI as sg 


def select_author_values(df: pd.DataFrame, author: str):
    _df = df.copy()
    _df = _df[_df["担当者"] == author]
    return _df[["新規", "機変", "クレカ", "でんき", "ネットワーク", "他"]].sum().to_frame().T
    
def get_year():
    files = os.listdir("成績管理表")
    arr = []
    for file in files:
        if ".ipynb_checkpoints" in file:
            continue
        arr.append(int(file.split("年")[0][-4:]))
    return arr 

    
def main():
    years = get_year()
    
    layout = [[sg.Text("過去のデータから売上数を合計します。")],
              [sg.Text("担当者を指定してください。")],
              [sg.Text("担当者名", size=(10, 1)), sg.InputText(key="-NAME-", default_text="(例) tanaka")], 
             [sg.Combo(years, key="-YEAR-", default_value=years[len(years)-1]), sg.Text("年")], 
             [sg.Combo(list(range(1, 13)), key="-MONTH-", default_value=1), sg.Text("月")], 
              [sg.Output(size=(80, 5))], 
             [sg.Button("実行", key="-SUBMIT-")]]
    
    window = sg.Window("売上確認アプリ", layout, size=(400, 400))
    
    while True:
        event, values = window.read()
        
        if event == sg.WIN_CLOSED:
            break 
        
        if event == "-SUBMIT-":
            name = values["-NAME-"]
            year = int(values["-YEAR-"])
            month = int(values["-MONTH-"])
            
            try:
            
                df = pd.read_excel(f"成績管理表/{str(year)}年分.xlsx", sheet_name=str(month) + "月")
                names = df["担当者"].unique().tolist()

                if(name == "" or name == "(例) tanaka" or name not in names):
                    print("担当者名が存在しません。")
                    time.sleep(2)
                    print("")
                else:
                    x = select_author_values(df, name)
                    shinki = x["新規"].values[0]
                    kihen = x["機変"].values[0]
                    card = x["クレカ"].values[0]
                    denki = x["でんき"].values[0]
                    nw = x["ネットワーク"].values[0]
                    other = x["他"].values[0]
                    print(f"新規: {shinki} 機変: {kihen} クレカ: {card} でんき: {denki} NW: {nw} 他: {other}")

            except Exception as e:
                print("予期せぬエラーが発生しました。")
                time.sleep(2)
                break 
            
    window.close()
    

if __name__ == "__main__":
    main()