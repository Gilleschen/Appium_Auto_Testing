# 使用說明
#### 使用前設定：

* 安裝 Appium (http://appium.io/)

* 環境設定 (請參考<a href="http://www.qa-knowhow.com/?p=2363">Appium手機自動化測試從頭學 –Windows/Android環境安裝篇</a>)

* 啟動 Appium Server (相關參數設定請參考<a href="http://www.automationtestinghub.com/appium-desktop-client-features/">Automation Testing Hub</a>)

* 下載<a href="https://github.com/Gilleschen/Appium_Auto_Testing_Android/raw/master/Appium_Android.jar"> Appium_Android.jar</a>及<a href="https://github.com/Gilleschen/Appium_Auto_Testing/blob/master/TestScript.xlsm"> TestScript.xlsm</a>

#### 測試腳本建立流程：

1. 於C:\建立 TUTK_QA_TestTool 資料夾 (C:\TUTK_QA_TestTool)

2. 於 TUTK_QA_TestTool 中建立TestTool資料夾與TestReport資料夾

3. 將 TestScript.xlsm 放至TestTool資料夾 (C:\TUTK_QA_TestTool\TestTool\TestScript.xlsm)(檔名及副檔名請勿更改)

4. 開啟 TestScript.xlsm 並允許啟動巨集 (已建立APP&Device、ExpectResult及說明工作表)

5. APP&Device工作表輸入APP Packageanme、APP Activity、測試裝置UDID、測試裝置OS版本、待測試腳本(以_TestScript結尾的工作表)、測試案例名稱、Appium_Android.jar路徑及Reset APP，範例如下圖：

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/APPAndDevice_3.PNG)

6. 建立測試腳本：新增一工作表，工作表名稱須以_TestScript為結尾 (e.g. Login_TestScript)，請參考[腳本產生器](#scriptcreater)，目前支援指令如下: (有區分大小寫，使用方式請參考TestScript.xlsm內說明工作表) 

          CaseName=>測試案列名稱(各案列開始時第一個填寫項目，必填!!!)

          Back=>點擊手機返回鍵

          Byid_Click/ByXpath_Click=>根據id/xpath搜尋元件並點擊元件

          Byid_LongPress/ByXpath_LongPress=>根據id/xpath搜尋元件並長按元件

          Byid_VerifyText/ByXpath_VerifyText=>根據id/xpath搜尋元件並取得元件Text屬性之字串後，比對ExpectResult內容

          Byid_SendKey/ByXpath_SendKey=>根據id/xpath搜尋元件並輸入數值或字串

          Byid_Clear/ByXpath_Clear=>根據id/xpath搜尋元件並清除數值或字串

          Byid_Wait/ByXpath_Wait=>根據id/xpath等待元件

          Byid_invisibility/ByXpath_invisibility=>根據id/xpath搜尋元件並等待該元件消失

          Byid_Swipe/ByXpath_Swipe=>根據id/xpath將元件A垂直移動到元件B位置,產生垂直滑動畫面

          ByXpath_Swipe_Vertical/ByXpath_Swipe_Horizontal=>垂直滑動/水平滑動n次

          Swipe=>根據x,y座標滑動畫面n次

          ByXpath_Swipe_FindText_Click_Android=>透過垂直/水平滑動畫面，點擊指定元件

          HideKeyboard=>關閉鍵盤

          Home=>點擊手機Home鍵

          LaunchAPP=>啟動APP&Device工作表指定的Packageanme之Activity

          Orientation=>切換手機Landscape及Portrait模式

          Power=>點擊手機電源鍵

          QuitAPP=>關閉APP&Device工作表指定的Packageanme之Activity

          ResetAPP=>重置APP(清除APP暫存紀錄)並重新啟動APP

          ScreenShot=>螢幕截圖

          Sleep=>閒置APP n秒鐘
  
範例腳本如下圖：

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testcase_example.PNG)
  
7. ExpectResult工作表：針對「字串」進行比對
   
   * A欄第二列處往下填入案列名稱 (CaseName)
        
   * 與案例名稱同列處輸入期望「字串」結果
        
 ExpectResult範例如下圖：
 
 ![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Result_example.PNG)

#### 測試腳本語法檢查：

1. 執行TestScript.xlsm增益集工具進行語法與資訊檢查，如下圖：

![image](https://github.com/Gilleschen/Android_invoke_excel/blob/master/picture/Gain_set.PNG)

2. 各功能說明：

        2.1 執行腳本：執行指定的工作表腳本，執行腳本前請確認以下4項功能無誤
        
        2.2 檢查資訊：確認APP&Device工作表所有欄位是否正確
        
        2.3 檢查案例語法：確認各案例結束後均執行QuitAPP方法
        
        2.4 檢查案例輸入值：確認所有命令及參數是否正確
        
        2.5 檢查期望結果：確認案例之期望字串是否列於ExpectResult工作表，當然非所有案列都需列ExpectResult    
        
        註：2.3, 2.4及2.5功能僅檢查以_TestScript為結尾且未隱藏的工作表 

3. 功能異常排除：針對功能無法正常運作

        3.1 移除增益集自訂工具列，如下圖：
        
      ![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/troubleshooting.png)
        
        3.2 存檔TestScript.xlsm
        
        3.3 重新開啟TestScript.xlsm

#### Excel 測試報告

1. 開啟 C:\TUTK_QA_TestTool\TestReport\TestReport.xlsm

2. 根據手機UDID自動建立TestReport工作表，如下圖： (e.g. abc123ABC123_TestReport)

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testreport_sheet_example.PNG)

範例測試結果如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/TestResult.PNG)

#### Log 紀錄

針對Error之測試案例，進行log紀錄，存放於 C:\TUTK_QA_TestTool\TestReport\\[APP Packagename]\\[Case Name]\\[Device UDID]\\log

<a name="scriptcreater"/>

#### 腳本產生器說明 

點擊增益集中腳本產生器，如下圖：

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ScriptCreator.png)

1. 指令類型按鈕(藍框)，列出指令清單(綠框)

2. 點選指令清單中的指令(綠框)後，透過Add按鈕加入右側的腳本清單(紫框)

3. 腳本清單完成後，點擊Create Case按鈕

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ScriptCreator3.png)

# 序列測試

1. 啟動與測試裝置相同數量的Appium (例如：要測試兩支裝置，則啟動兩組Appium)

2. 進入Advanced欄位 (設定Server Address, Server Port)

   2.1 固定Server Address = 127.0.0.1

   2.3 第一組Server Port設定為4723；第二組Server Port設定為4725，如下圖 (若有第三組Appium，則Server Port設定為4727(即4725+2)，每次port都+2，依此類推)

第一組Server：
![image](https://github.com/Gilleschen/Appium_Auto_Testing/blob/master/picture/serverone.png)

第二組Server：
![image](https://github.com/Gilleschen/Appium_Auto_Testing/blob/master/picture/servertwo.png)

3. 啟動各組Appium Server

#### 備註：

* Appium Client Libraries Version: java-client-6.1.0

* Selenium Client Version: 3.14.0

* Excel欄位若輸入純數字(e.g. 8888)，請轉換為文字格式，皆於數字前面加入單引號 (e.g. '8888)或執行增益集的檢查案例輸入值功能

* 固定Server Address = 127.0.0.1, 預設Server Port = 4723

* Appium NEW_COMMAND_TIMEOUT=120 second ;WebDriverWait timeout=30 second

* *目前不支援WiFi指令*


