# 使用說明

#### Framework

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/Framework.png)

#### 使用前設定

* 安裝 Appium (http://appium.io/)

* 環境設定 (請參考<a href="http://www.qa-knowhow.com/?p=2363">Appium手機自動化測試從頭學 –Windows/Android環境安裝篇</a>)

* 啟動 Appium Server (相關參數設定請參考<a href="http://www.automationtestinghub.com/appium-desktop-client-features/">Automation Testing Hub</a>)

* 下載<a href="https://github.com/Gilleschen/Appium_Auto_Testing_Android/raw/master/Appium_Android.jar"> Appium_Android.jar</a>及<a href="https://github.com/Gilleschen/Appium_Auto_Testing/blob/master/TestScript.xlsm"> TestScript.xlsm</a>至C:\TUTK_QA_TestTool\TestTool資料夾

* 建立C:\TUTK_QA_TestTool\TestReport資料夾

#### 腳本建立流程

Step 1. 開啟TestScript.xlsm 並允許啟動巨集 (已建立APP&Device、APP&Device_Data、ExpectResult及說明工作表)

Step 2. 建立腳本：新增一工作表，工作表名稱必需以_TestScript為結尾 (如：Login_TestScript)，可透過[腳本產生器](#scriptcreater)建立腳本，指令(區分大小寫)如下：(指令使用方式請參考TestScript.xlsm內說明工作表) 

   * CaseName: 測試案列名稱(各案列開始時第一個填寫項目，必填!!!)
    
   * Back: 點擊行動裝置返回鍵

   * Byid_Click/ByXpath_Click: 根據id/xpath搜尋元件並點擊元件

   * Byid_LongPress/ByXpath_LongPress: 根據id/xpath搜尋元件並長按元件
    
   * Byid_VerifyText/ByXpath_VerifyText: 根據id/xpath搜尋元件並取得元件Text屬性之字串後，比對ExpectResult內期望字串

   * Byid_SendKey/ByXpath_SendKey: 根據id/xpath搜尋元件並輸入字串

   * Byid_Clear/ByXpath_Clear: 根據id/xpath搜尋元件並清除字串

   * Byid_Wait/ByXpath_Wait: 根據id/xpath等待元件出現

   * Byid_invisibility/ByXpath_invisibility: 根據id/xpath搜尋元件並等待該元件消失

   * Byid_Swipe/ByXpath_Swipe: 根據id/xpath將元件A移動到元件B位置，產生滑動畫面效果

   * ByXpath_Swipe_Vertical/ByXpath_Swipe_Horizontal: 垂直滑動/水平滑動

   * Swipe: 根據x,y座標滑動畫面

   * ByXpath_Swipe_FindText_Click_Android: 透過垂直/水平滑動畫面，點擊指定元件

   * HideKeyboard: 關閉鍵盤

   * Home: 點擊行動裝置Home鍵

   * LaunchAPP: 啟動APP&Device工作表指定的PackageName之Activity

   * Orientation: 切換行動裝置Landscape及Portrait模式

   * Power: 點擊行動裝置電源鍵

   * QuitAPP: 關閉APP&Device工作表指定的PackageName
    
   * ResetAPP: 重置APP(清除APP暫存紀錄)並重新啟動APP

   * ScreenShot: 螢幕截圖

   * Sleep: 閒置APP
  
範例腳本如下圖：

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/%E6%A1%88%E4%BE%8B%E8%AA%AA%E6%98%8E.png)
 
Step 3. 設定「期望字串」：點擊ExpectResult工作表，當使用Byid_VerifyText或ByXpath_VerifyText時(如上圖範例#7)，需在ExpectResult工作表填入期望字串，可藉由[編輯期望字串](#ExpectResult)功能進行新增。 (若測試案例不包含檢驗字串，則此步驟可省略)
   
   * A欄第二列處往下填入案列名稱 (CaseName)
        
   * 與案例名稱同列處輸入「期望字串」
        
範例如下圖：
 
 ![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ExpectResult_example.PNG)
 
 Step 4. 設定測試APP、測試裝置及測試腳本等資訊：點擊APP&Device工作表，可藉由[APP、設備與腳本](#testInformations)進行設定，項目如下：

* APP Package name與Activity
   
* 測試裝置UDID與OS版本
   
* 腳本工作表名稱與測試案例名稱 (若不指定案例，則執行所有案例)
   
* 測試前是否Reset APP

範例如下圖：

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/APPAndDevice_3.PNG)
 
Step 5. 點擊執行腳本，如下圖：
 
 ![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/RunScript.png)
 
 #### Excel 測試報告

Step 1. 開啟 C:\TUTK_QA_TestTool\TestReport\TestReport.xlsm

Step 2. 根據行動裝置UDID自動建立TestReport工作表，如下圖： (e.g. abc123ABC123_TestReport)

![image](https://github.com/Gilleschen/APP_Vsaas_2.0_Android_invoke_excel_Result_try_catch/blob/master/picture/Testreport_sheet_example.PNG)

範例測試結果如下圖：

![image](https://github.com/Gilleschen/Web_Auto_Testing/blob/master/picture/TestResult.PNG)

#### Log 紀錄

針對Error之測試案例，進行log紀錄，存放於 C:\TUTK_QA_TestTool\TestReport\\[APP Packagename]\\[Case Name]\\[Device UDID]\\log

#### VBA 巨集

1. 點擊增益集工具，如下圖 (11項功能)：

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/functions.PNG)

2. 各功能說明：

* 執行腳本：開始測試。
        
* 檢查APP與裝置資訊：確認APP&Device工作表所有欄位是否填寫。
        
* 檢查案例語法：確認各案例結束後均執行QuitAPP方法。
        
* 檢查案例參數：確認所有指令及參數是否正確。
        
* 檢查期望字串：確認期望字串是否填入ExpectResult工作表。
        
* [APP、設備與腳本](#testInformations)：透過VBA設定待測試的APP、UDID及腳本案例，也可手動填寫APP&Device工作表。
        
* [管理UDID/PackageName](#APPandDevice)：新增常用的Package name或UDID，也可手動填寫APP&Device_Data工作表。
        
* [腳本產生器](#scriptcreater)：透過VBA建立新腳本，也可手動建立工作表腳本。
        
* [腳本編輯器](#scripteditor)：透過VBA編輯現有腳本，也可手動編輯工作表腳本。
        
* [腳本檢查](#commandcheck)：針對指定的腳本進行指令與參數檢查。
        
* [編輯期望字串](#ExpectResult)：透過VBA編輯/新增期望字串，也可手動編輯ExpectResult工作表字串。

# VBA 巨集使用說明

<a name="testInformations"/>

#### APP、設備與腳本說明

Step 1. 選擇Package name後，自動列出對應的Activity (藍框)

Step 2. 選擇Udid後，自動列出對應的OS版本 (綠框)

Step 3. 選擇測試腳本，自動列出該腳本下的測試案例名稱 (橘框)

Step 4. 設定測試前是否重置APP

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/TestInformations.png)

<a name="APPandDevice"/>

#### 管理UDID/PackageName說明

Step 1. 選擇Device UDID後(綠框)，列出常用的UDID及OS版號

Step 2. 選擇APP Packagename後 (藍框)，列出儲存的Package name 及 Activity

Step 3. 編輯文字方塊，點擊Add/Edit進行編輯

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/EditUdidandPackagename.png)

備註：常用的Package name、Activity、測試裝置UDID與OS版本儲存於APP&Device_Data工作表如下圖。

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/APPandDevice_Data.PNG)

<a name="scriptcreater"/>

#### 腳本產生器說明 

Step 1. 點擊指令類型按鈕(藍框)，列出指令清單(綠框)

Step 2. 點選指令清單中的指令(綠框)後，點擊Add按鈕加入右側的腳本清單(紫框)

Step 3. 腳本完成後，點擊Create Case按鈕

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ScriptCreator3.png)

<a name="scripteditor"/>

#### 腳本編輯器說明

Step 1. 選擇腳本 (綠框)

Step 2. 選擇腳本中案例 (藍框)，列出該案例中的指令 (橘框)

Step 3. 點擊指令類型按鈕

Step 4. 選擇指令，點擊Add/Up/Down/Delete進行編輯

Step 5. 腳本編輯完成後，點擊Edit Case按鈕

備註：若想新增新案例，可勾選New Case (紫框)。

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ScriptEditor.png)

<a name="commandcheck"/>

#### 腳本檢查說明

Step 1. 選擇檢查項目 (綠框)

Step 2. 選擇腳本 (藍框)

Step 3. 點擊Add加入腳本 (橘框)

Step 4. 點擊Start

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/commandCheck.png)

<a name="ExpectResult"/>

#### 編輯期望字串說明

Step 1. 選擇案例 (綠框)

Step 2. 選擇字串 (藍框)

Step 3. 修改或建立字串後 (紫框)，點擊Add/Edit String 加入String list中

備註：點擊Add Case可建立新案例至ExpectResult工作表。

![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/ExpectResultCreator.png)

#### VBA 功能異常排除

Step 1. 移除增益集自訂工具列，如下圖：
        
![image](https://github.com/Gilleschen/Appium_Auto_Testing_Android/blob/master/picture/troubleshooting.png)
        
Step 2. 存檔TestScript.xlsm
        
Step 3. 重新開啟TestScript.xlsm


# 序列測試

Step 1. 啟動與測試裝置相同數量的Appium (例如：要測試兩支裝置，則啟動兩組Appium)

Step 2. 進入Advanced欄位 (設定Server Address, Server Port)：

* 固定Server Address = 127.0.0.1

* 第一組Server Port設定為4723；第二組Server Port設定為4725，如下圖 (若有第三組Appium，則Server Port設定為4727(即4725+2)，每次port      都+2，依此類推)

第一組Server：
![image](https://github.com/Gilleschen/Appium_Auto_Testing/blob/master/picture/serverone.png)

第二組Server：
![image](https://github.com/Gilleschen/Appium_Auto_Testing/blob/master/picture/servertwo.png)

Step 3. 啟動各組Appium Server

#### 備註

* Appium Client Libraries Version: java-client-6.1.0

* Selenium Client Version: 3.14.0

* Excel欄位若輸入純數字(e.g. 8888)，請轉換為文字格式，皆於數字前面加入單引號 (e.g. '8888)或執行增益集的檢查案例輸入值功能

* 固定Server Address = 127.0.0.1, 預設Server Port = 4723

* Appium NEW_COMMAND_TIMEOUT=120 second ;WebDriverWait timeout=30 second

* *目前不支援WiFi指令*


