# KMU_wac_webcrawler
高醫wac系統選課中籤率爬蟲
我懶得開wac於是寫了爬蟲幫忙直接抓選課中籤率
# Requirement
1.	本程式含有python. selenium. openpyxl套件
2.	selenium需要webdriver(你需要對應版本的瀏覽器)
3.	程式內有路徑需要自己修正
4.	我登陸wac的密碼已經改掉了 你找不到XD
備註: 如果你是windows系統可以用windows工作排程設定固定時間執行
# Intro
1.	GGG.xlsx為爬蟲完儲存資料的檔案
含有多個worksheet(不同時間的中籤率)
2.	wacwc.py為主程式 負責登錄wac=>爬蟲翻頁=>將資料存在GGG.xlsx
3.	chromedriver.exe為wacwc.py運行時需要的google chrome瀏覽器啟動器
4.	GG.bat可用可不用 單純執行wacwc.py 使用.bat檔案目的是可以讓windows電腦自動排程
# Future
未來會把我抓的資料做成圖片解析高醫學生選課的集中時間(折線圖等)
會把CSV. EXCEL檔放到雲端或匯入網頁供大家服用

