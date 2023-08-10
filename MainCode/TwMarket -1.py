from bs4 import BeautifulSoup
import requests
import lxml
import pandas as pd
import pygsheets
import datetime


### 1.資訊來源連結 ###
# 上市三大法人買賣超
twse_url = "https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=html"
# 期貨三大法人
Taifex_futures_url = "https://www.taifex.com.tw/cht/3/futContractsDateExcel"
# 選擇權外資與自營商
Taifex_OP_url = "https://www.taifex.com.tw/cht/3/callsAndPutsDateExcel"

### 2.建立GoogleSheet 連線 ###
#建立金鑰連結
KeyPath = r"放自己的路徑"
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

# 期貨三大法人
def Taifex_futues(Taifex_futures_url):
    Taifex_df = pd.read_html(Taifex_futures_url)[0]
    new_futures_df = Taifex_df.iloc[[4,5,6,7],[2,3,5,7,9,11,13]]
    Taifex_df_reset = new_futures_df.drop(index=[4])
    Taifex_df_reset.columns = ['身份別',
                           '多方交易','空方交易','交易口數差額',
                           '多方未平倉','空方未平倉','未平倉口數差額']
    Taifex_df_reset = Taifex_df_reset.set_index(['身份別'])
    return Taifex_df_reset
Taifex_futues_data = Taifex_futues(Taifex_futures_url)

#選擇權三大法人
def Taifex_op(Taifex_OP_url):
    Taifex_OP_data = pd.read_html(Taifex_OP_url)[0]
    new_Taifex_data = Taifex_OP_data.iloc[[3,5,7,8,10],[2,3,4,6,8,10,12,14]]
    new_Taifex_data.columns = ['權別',
                             '身份別',
                             '買方(口數)','賣方(口數)','交易口數差額',
                             '買方(口數)','賣方(口數)','未平倉口數差額']
    new_Taifex_data_1 = new_Taifex_data.drop(index=[3])
    new_Taifex_data_1 = new_Taifex_data_1.set_index(['權別'])
    return new_Taifex_data_1
Taifex_op_data = Taifex_op(Taifex_OP_url)
    

### 4.更新時間 ###
update_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

### 5.資料寫入指定表單與工作表
WriteSheet_Twse = sh.worksheet_by_title("大盤資訊")
WriteSheet_Twse.clear()

#寫入上市資料
WriteSheet_Twse.update_value('A1','大盤三大法人買賣超 資料更新時間 : '+update_time)
WriteSheet_Twse.set_dataframe(Twse_data, 'A2', copy_index=True, nan='')

#寫入期貨資料
WriteSheet_Twse.update_value('A10','大型台指期貨三大法人資訊')
WriteSheet_Twse.set_dataframe(Taifex_futues_data, 'A11', copy_index=True, nan='')

#寫入選擇權資料
WriteSheet_Twse.update_value('A16','台指選擇權自營商&外資')
WriteSheet_Twse.set_dataframe(Taifex_op_data, 'A17', copy_index=True, nan='')