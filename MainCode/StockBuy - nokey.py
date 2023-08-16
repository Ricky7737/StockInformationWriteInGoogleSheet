import pandas as pd
import pygsheets
import datetime

### 1.資訊來源連結 ###
#上市資料
#外資買賣超資料
TW_ForInv_url = 'https://www.twse.com.tw/rwd/zh/fund/TWT38U?response=html'
#投信買賣超連結
Tw_IvenBank_url = 'https://www.twse.com.tw/rwd/zh/fund/TWT44U?response=html'

#上櫃資料
update_time_roc = str(int(datetime.datetime.now().strftime('%Y'))-1911)+datetime.datetime.now().strftime('/%m/%d')
#外資買超資料
OTC_ForInv_url = 'https://www.tpex.org.tw/web/stock/3insti/qfii_trading/forgtr_result.php?l=zh-tw&t=D&type=buy&d='+update_time_roc+'&s=0,asc&o=htm'
#投信買超資料
OTC_IvenBank_url = 'https://www.tpex.org.tw/web/stock/3insti/sitc_trading/sitctr_result.php?l=zh-tw&t=D&type=buy&d='+update_time_roc+'&o=htm'

### 2.處理資料 ###
#處理上是與上櫃外資資料
def DelForData30(url):
    ForData30 = pd.read_html(url)[0]
    ForData30 = ForData30.iloc[:30, [1,2,11]]
    ForData30.columns = ['證券代號','證券名稱','買超張數']
    ForData30 = ForData30.set_index(['證券代號'])
    return ForData30

TW_ForInv30_df = DelForData30(TW_ForInv_url)
#這邊處理一下上市的買超，資料其實是股數，所以特別處理
TW_ForInv30_df['買超張數'] = TW_ForInv30_df['買超張數']//1000
OTC_ForInv30_df = DelForData30(OTC_ForInv_url) 

#處理上市與上櫃投信資料
def DelInvesBank20(url_1):
    IBData20 = pd.read_html(url_1)[0]
    IBData20 = IBData20.iloc[:20, [1,2,5]]
    IBData20.columns = ['證券代號','證券名稱','買超張數']
    IBData20 = IBData20.set_index(['證券代號'])
    return IBData20

Tw_IvenBank20_df = DelInvesBank20(Tw_IvenBank_url)
Tw_IvenBank20_df['買超張數'] = Tw_IvenBank20_df['買超張數']//1000
OTC_IvenBank20_df = DelInvesBank20(OTC_IvenBank_url)

### 3. 處理GoogleSheet 資料 ###
#建立金鑰連結
KeyPath = r"放入自己的.json"
gc = pygsheets.authorize(service_account_file=KeyPath)
#指定 google表單寫入的連結
googlesheet_url = "放入自己的googlesheet連結"
sh = gc.open_by_url(googlesheet_url)

## 指定寫入的資料報表
#先做表格清理
WriteSheet_TWSE = sh.worksheet_by_title("上市個股買賣超排名")
WriteSheet_TWSE.clear()
WriteSheet_OTC = sh.worksheet_by_title("上櫃個股買賣超排名")
WriteSheet_OTC.clear()

##寫入資料欄位大標題
#上市標題
WriteSheet_TWSE.update_value('A1','上市外資與投信個股買賣超排名 資料更新時間 : '+update_time_roc)
#上櫃標題
WriteSheet_OTC.update_value('A1','上櫃外資與投信個股買賣超排名 資料更新時間 : '+update_time_roc)

##寫入上市資料
#寫入外資買超前 50
WriteSheet_TWSE.update_value('A2','外資買超前 30')
WriteSheet_TWSE.set_dataframe(TW_ForInv30_df, 'A3', copy_index=True, nan='')

#寫入投信買超前 20
WriteSheet_TWSE.update_value('E2','投信買超前 20')
WriteSheet_TWSE.set_dataframe(Tw_IvenBank20_df, 'E3', copy_index=True, nan='')

##寫入上櫃資料
#寫入外資買超前 50
WriteSheet_OTC.update_value('A2','外資買超前 30')
WriteSheet_OTC.set_dataframe(OTC_ForInv30_df, 'A3', copy_index=True, nan='')

#寫入投信買超前 20
WriteSheet_OTC.update_value('E2','投信買超前 20')
WriteSheet_OTC.set_dataframe(OTC_IvenBank20_df, 'E3', copy_index=True, nan='')

