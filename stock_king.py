import datetime
# import numpy as np
import pandas as pd
# import matplotlib
#專門做『技術分析』的套件
from talib import abstract
#專門抓台股的套件
# import twstock
# import base64
from email.mime.text import MIMEText

#製作 Stock_ID
st_data = pd.read_excel('stock_id_00.xlsx',encoding='gbk')
ind=[str(i) for i in list(st_data.num)]
s_name=list(st_data.name)
st_d={}
for i in range(len(ind)):
    st_d[ind[i]]=s_name[i]
    
stock_id=pd.read_excel('stock_id.xlsx',dtype=str)
s_id=list((stock_id.iloc[0])) #全部的股票代號
#已買清單
buy_list=list(pd.read_excel('buy_list.xlsx',dtype=str)[0])
##############################################
#過去的推薦買賣股票
try:
    b=pd.read_excel('buy_id.xlsx',dtype={0:str})[0]
except:
    b=[]
try:
    s=pd.read_excel('sell_id.xlsx',dtype={0:str})[0]
except:
    s=[]
buy_last=list(b[:])
sell_last=list(s[:])

sum_buy=0
sum_sell=0
################################################
#設定爬蟲股票代號
# s_id=['0050']
buy=[]
sell=[]

for i in s_id:
    sid = i
    #設定爬蟲時間
    start = datetime.datetime.now() - datetime.timedelta(days=360)
    end = datetime.date.today()
    #導入pandas_datareader
    from pandas_datareader import data
    # 與yahoo請求，套件路徑因版本不同
    pd.core.common.is_list_like = pd.api.types.is_list_like
    try:
        stock_dr = data.get_data_yahoo(sid+'.TW', start, end, interval='w')
    except:
        print(i+" 已經下市")
        continue
    stock_dr.columns=['high','low','open','close','volume','adj close']
    ####################################################################
    if stock_dr.iloc[-1].open < 300 and stock_dr.iloc[-1].volume>5000000:  #取價格小於 50000,且量大於500ㄋ 
        #accuracy
        if i in buy_last:
            if stock_dr.iloc[-1].open <  stock_dr.iloc[-1].close:
                sum_buy+=1
        if i in sell_last:
            if stock_dr.iloc[-1].open >  stock_dr.iloc[-1].close:
                sum_sell+=1        
        
        #收盤價
        #df['close'].plot(figsize=(16, 8))
        # KD
        # kd=abstract.STOCH(stock_dr)
        # abstract.STOCH(stock_dr).plot(figsize=(16, 8))
        #MACD

        kd=abstract.STOCH(stock_dr,fastk=9)
        #畫圖
        #abstract.MACD(stock_dr).plot(figsize=(16, 8)) 
        #stock_dr['close'].plot(secondary_y=True)
        cross=kd.iloc[-3:]
        #找交叉
        for j in range(len(cross)-1):
            #print(j)
            if cross.slowd.iloc[j] > cross.slowk.iloc[j] and cross.slowd.iloc[j+1] < cross.slowk.iloc[j+1] and cross.slowd.iloc[j+1]<30:
                #print(' +  : ',cross.index[j])
                buy+=[i,st_d[i],stock_dr.iloc[-1].open,cross.index[j],' + ']
            elif cross.slowd.iloc[j] < cross.slowk.iloc[j] and cross.slowd.iloc[j+1] > cross.slowk.iloc[j+1] :
                
                #print(' -  : ',cross.index[j])
                if i in buy_list:
                    sell+=[i,st_d[i],stock_dr.iloc[-1].open,cross.index[j],' - ***']
                else:
                    sell+=[i,st_d[i],stock_dr.iloc[-1].open,cross.index[j],' - ']
#print('buy : ',buy)
#print('sell : ',sell)
buy_id=[]
sell_id=[]
for i in range(len(buy)):
    if i % 5==0:
        buy_id+=[[buy[i],buy[i+1],buy[i+2]]]
for i in range(len(sell)):
    if i % 5==0:
        sell_id+=[[sell[i],sell[i+1],sell[i+2]]]
print('buy : ',buy_id)
print('\n')
print('sell : ',sell_id)
print('#################################################################################')
print('Accuracy')
try:
    print('Buy : ',sum_buy/len(buy_last))
    print('Sell : ',sum_sell/len(sell_last))
except:
    print("尚未有買賣紀錄")
buy_id_pd=pd.DataFrame(buy_id)
sell_id_pd=pd.DataFrame(sell_id)
buy_id_pd.to_excel('buy_id.xlsx',index=False)
sell_id_pd.to_excel('sell_id.xlsx',index=False)

##############################################################
#寄信
import smtplib
smtp=smtplib.SMTP('smtp.gmail.com', 587)
smtp.ehlo()
smtp.starttls()
smtp.login('hsustock12345@gmail.com','Qaz78900')

#寄/收件人
from_addr='hsustock12345@gmail.com'
to_addr=['hsusean1219@gmail.com','stevenlinlyc860415@gmail.com','spencer8005@yahoo.com.tw','davidlv7621@yahoo.com.tw','anderson831208@gmail.com']
# to_addr=['hsusean1219@gmail.com','Jeremy.Hsu@sti.com.tw']
# to_addr=['hsusean1219@gmail.com']

#推薦名單
recommend_buy = buy_id
recommend_sell = sell_id

#編輯內文
msg=""
msg+="Buy\n"
if len(recommend_buy)==0:
    msg+="None!!\n\n"
else:
    for i in recommend_buy:
        if i == recommend_buy[-1]:
            msg+= i[0]+' '+i[1]+" Price "+ str(round(i[2],2))+" !\n\n"
        else:
            msg+= i[0]+' '+i[1]+" Price "+ str(round(i[2],2))+",\n"
msg+="Sell\n"
if len(recommend_sell)==0:
    msg+="None!!\n\n"
else:
    for i in recommend_sell:
        if i == recommend_sell[-1]:
            msg+= i[0]+' '+i[1]+" Price "+ str(round(i[2],2))+" !\n\n"
        else:
            msg+= i[0]+' '+i[1]+" Price "+ str(round(i[2],2))+",\n"
            
#輸入內容
text = MIMEText(msg, 'plain', 'utf-8')
text['From'] = u'台灣巴菲哲'
text['Subject'] =u'自動報明牌系統'

#寄信        
for k in to_addr:
    status=smtp.sendmail(from_addr, k, text.as_string())#加密文件，避免私密信息被截取 發現信的內容不能有":"            

#確認
if status=={}:
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+"  郵件傳送成功!")
else:
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+"  郵件傳送失敗!")
smtp.quit()