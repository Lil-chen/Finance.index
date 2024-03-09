import requests
from bs4 import BeautifulSoup
import pandas as pd 

#我們主要利用invest.com網站的資料當作我們的總經指標的來源
urls = ["https://www.investing.com/indices/taiwan-weighted-historical-data", 
        "https://www.investing.com/indices/us-30-historical-data",
        "https://www.investing.com/indices/us-spx-500-historical-data",
        "https://www.investing.com/indices/nasdaq-composite-historical-data",
        "https://www.investing.com/indices/shanghai-composite-historical-data",
        "https://www.investing.com/indices/hang-sen-40-historical-data",
        "https://www.investing.com/indices/japan-ni225-historical-data",
        "https://www.investing.com/currencies/usd-twd-historical-data",
        "https://www.investing.com/currencies/usd-cny-historical-data",
        "https://www.investing.com/currencies/usd-jpy-historical-data",
        "https://www.investing.com/currencies/usd-eur-historical-data",
        "https://www.investing.com/currencies/usd-gbp-historical-data",
        "https://www.investing.com/commodities/brent-oil-historical-data",
        "https://www.investing.com/commodities/crude-oil-historical-data",
        "https://www.investing.com/commodities/natural-gas-historical-data",
        "https://www.investing.com/commodities/gold-historical-data",
        "https://www.investing.com/commodities/us-soybeans-historical-data",
        "https://www.investing.com/commodities/us-wheat-historical-data",
        "https://www.investing.com/commodities/us-corn-historical-data",
        "https://www.investing.com/commodities/us-cocoa-historical-data",
        "https://www.investing.com/commodities/us-cotton-no.2-historical-data",
        "https://www.investing.com/indices/baltic-dry-historical-data"
        ]
dfs = []
lists=[]
list_year=[]
listsss=[]

for url in urls:
    response = requests.get(url)
    html = response.text
    soup = BeautifulSoup(html, "html.parser")

    # 收盤價(Now)
    closed = soup.findAll("div", attrs={"data-test": "instrument-price-last"})
    for price in closed:
        lists.append(price.text)
    
    #diff
    diffs=soup.findAll("span", attrs={"data-test": "instrument-price-change-percent"})
    for diff in diffs:
        listsss.append(diff.text)
    
    listss = []

    #High Low 
    another_price = soup.findAll("td", attrs={"data-test": "relative-most-active-last"})
    #根據網站 我們抓前15筆資料 剛好是5天的open high low
    if len(another_price) >= 3:
        listss.extend([price.text for price in another_price[:15]])

        
        df = pd.DataFrame({
            "High": listss[1::3],
            "Low": listss[2::3]
        })

        # 每五天資料的最大值最小值，當作week low week high  
        max_high = df["High"].max()
        min_low = df["Low"].min()

        
        total_df = pd.DataFrame({
            "High": [max_high],
            "Low": [min_low]
        })

        
        dfs.append(total_df)

    else:
        print("Error: Insufficient data in another_price")
    
#year high low 
    years_range = soup.findAll("div", attrs={"flex gap-2 font-bold tracking-[0.2px] items-center"})
    
    
    for div in years_range:
        span = div.findAll("span")
        for i in span:   
            list_year.append(i.text)
            
data = [(price) for price in list_year]
columns = ['Daily Low', 'Daily High', 'Year Low', 'Year High']
year_data = [data[i:i + 4] for i in range(0, len(data), 4)]
df2 = pd.DataFrame(year_data, columns=columns)

df = pd.concat(dfs, ignore_index=True)     
#我們把爬到的各項資料，用pandas做成一個dataframe   
name=["台灣加權股價指數","道瓊","sp500","那斯達克指數","上證指數","香港恒生指數","日經指數","美元/台幣","美元/人民幣","美元/日幣","美元/歐元","美元/英鎊","布蘭特原油","西德州原油","天然氣","黃金","黃豆","小麥","玉米","可可豆","棉花","BDI"]
df["stock"]=name
df["Now"]=lists
df["diff"]=listsss
df["High"]=df["High"]
df["Low"]=df["Low"]
df["Year High"]=df2["Year High"]
df["Year Low"]=df2["Year Low"]
for object in [ "Now", "High", "Low","Year High","Year Low"]:
    df[object]=df[object].str.replace(",","",regex=True).astype("float")
#用正則去括號
df["diff"] = df["diff"].replace(r'\((.*?)\)', r'\1', regex=True)

df = df.map(lambda x: "{:.2f}".format(x) if isinstance(x, (float, int)) else x)
df.set_index("stock", inplace=True)
df=df[[ "Now","diff","High","Low","Year High","Year Low"]]
print(df.head())
#完成
df.to_excel("財經周報1.xlsx")





