import datetime
import pandas as pd
from email.mime.text import MIMEText
import concurrent.futures
import smtplib
from KD_check import stock_m
#製作 Stock_ID
st_data = pd.read_excel('stock_id.xlsx',dtype=str)
s_id=[str(i) for i in list(st_data.num)] #全部的股票代號
s_name=list(st_data.name)
st_d={}
for i in range(len(s_id)):
    st_d[s_id[i]]=s_name[i]

##################################################################
#切割任務
#製作切割開始結束的位置list
div_list=[0]
part=120 #建多少 thread 一起跑
start_div=0
for i in range(part-1):
    div_list+=[start_div + len(s_id)//part]
    start_div+=len(s_id)//part
div_list+=[len(s_id)]
t=[]
buy=[]
sell=[]
out_of_market=[]
#平行處理
with concurrent.futures.ThreadPoolExecutor() as executor:
    for i in range(part):
        t+=[executor.submit(stock_m, s_id, st_d, div_list[i], div_list[i+1])]
    for i in range(part):
        if t[i].result()[0] != []:
            buy+=[t[i].result()[0][0]] 
        if t[i].result()[1] != []:
            sell+=[t[i].result()[1][0]] 
        if t[i].result()[2] != []:
            out_of_market+=[t[i].result()[2][0]]
    



buy_id=[]
sell_id=[]
for i in buy:
    buy_id+=[[i[0],i[1],i[2]]]
for i in sell:
    sell_id+=[[i[0],i[1],i[2]]]
print('buy : ',buy_id)
print('\n')
print('sell : ',sell_id)
print('#################################################################################')
##############################################################
#寄信

smtp=smtplib.SMTP('smtp.gmail.com', 587)
smtp.ehlo()
smtp.starttls()
smtp.login('hsustock12345@gmail.com','Qaz78900')

#寄/收件人
from_addr='hsustock12345@gmail.com'
to_addr=['hsusean1219@gmail.com','stevenlinlyc860415@gmail.com','spencer8005@yahoo.com.tw',
         'davidlv7621@yahoo.com.tw','anderson831208@gmail.com','Jeremy.Hsu@sti.com.tw','emailirene2006@gmail.com']
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
