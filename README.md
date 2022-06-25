# dpr
配当性向と目標配当性向の差分を分析する

```
pip install selenium  
pip install openpyxl  
pip install urllib3  
pip install tika  
```

クロームドライバーをダウンロード  
[https://chromedriver.chromium.org/downloads]

上場企業一覧をダウンロード　使用する前にxls→xlsx変換を行うこと  
[https://www.jpx.co.jp/markets/statistics-equities/misc/01.html]

配当(fy-stock-dividend.csv)をダウンロード  
[https://irbank.net/download]


実行方法  
```
mkdir pdf
python download.py [途中から再実行する場合には証券コードを引数に]
python dpr.py
```