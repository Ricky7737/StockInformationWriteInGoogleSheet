from bs4 import BeautifulSoup
import requests
import lxml
import pandas as pd
import pygsheets
import datetime


### 1.資訊來源連結 ###
# 上市三大法人買賣超
twse_url = "https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=html"

### 2.建立GoogleSheet 連線 ###
#建立金鑰連結
KeyPath = r" 放自己的googleAPI .json路徑"
gc = pygsheets.authorize(service_account_file=KeyPath)
#指定 google表單寫入的連結
googlesheet_url = "https://docs.google.com/spreadsheets/d/1BEw9S1YL-AugmGNGMA3P9mPCz4fRRrPzHYMz9dkBE3k/edit?usp=sharing"
sh = gc.open_by_url(googlesheet_url)

### 3. 處理不同資料的格式 ###
# 上市法人買賣超
def Tw_infomation(twse_url):
    twse_df_new = pd.read_html(twse_url)[0]
    twse_df_new.columns = ["單位名稱",
                           "買進金額", 
                           "賣出金額", 
                           "買賣差額"]
    twse_df_reset = twse_df_new.set_index('單位名稱')
    return twse_df_reset
Twse_data = Tw_infomation(twse_url)

### 4.更新時間 ###
update_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

### 5.資料寫入指定表單與工作表
WriteSheet_Twse = sh.worksheet_by_title("大盤資訊")
WriteSheet_Twse.clear()
WriteSheet_Twse.update_value('A1','大盤三大法人買賣超 資料更新時間 : '+update_time)

# 寫入上市資料
WriteSheet_Twse.set_dataframe(Twse_data, 'A2', copy_index=True, nan='')