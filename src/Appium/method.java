package Appium;

import static java.time.Duration.ofSeconds;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.ScreenOrientation;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.logging.LogEntry;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.base.Stopwatch;
import io.appium.java_client.TouchAction;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidKeyCode;

import io.appium.java_client.remote.AndroidMobileCapabilityType;
import io.appium.java_client.remote.MobileCapabilityType;
import io.appium.java_client.touch.LongPressOptions;
import io.appium.java_client.touch.WaitOptions;
import io.appium.java_client.touch.offset.ElementOption;
import io.appium.java_client.touch.offset.PointOption;

public class method {
	// int systemPort = 8200;// System port
	static int port = 4723;// Appium port
	static int StartPort = port;
	static int device_timeout = 1200;// 120sec
	int command_timeout = 30;// 30sec
	LoadExpectResult ExpectResult = new LoadExpectResult();
	static LoadTestCase TestCase = new LoadTestCase();
	static ArrayList<AndroidDriver> driver = new ArrayList<AndroidDriver>();// 儲存
																			// AndroidDriver
																			// 的ArrayList
	// driverBK備份driver 的ArrayList，目的為回復原本driver
	// list中的driver資訊，因為Command出錯後會設定原driver list為null，故當執行下個案例前需設定回原driver 資訊
	static ArrayList<AndroidDriver> driverBK = new ArrayList<AndroidDriver>();
	static String CaseErrorList[][] = new String[TestCase.CaseList.size()][TestCase.DeviceInformation.deviceName
			.size()];// 紀錄各案例於各裝置之指令結果 (2維陣列)CaseErrorList[CaseList][Devices]
	static AndroidDriver Androiddriver = null;
	static String ErrorList[] = new String[TestCase.DeviceInformation.deviceName.size()];// 紀錄各裝置之指令結果
	WebDriverWait[] wait = new WebDriverWait[TestCase.DeviceInformation.deviceName.size()];
	static XSSFWorkbook workBook;
	static String appElemnt;// APP元件名稱
	static String appInput;// 輸入值
	static String appInputXpath;// 輸入值的Xpath格式
	static String toElemnt;// APP元件名稱
	static int startx, starty, endx, endy;// Swipe移動座標
	static int iterative;// 畫面滑動次數
	static String scroll;// 畫面捲動方向
	static String appElemntarray;// 搜尋的多筆類似元件
	static String checkVerifyText;// 確認執行QuitAPP前是否執行過VerifyText
	static String switchWiFi;// 啟動wifi或關閉wifi
	String element[] = new String[driver.size()];
	static int CurrentCaseNumber = -1;// 目前執行到第幾個測試案列
	static Boolean CommandError = true;// 判定執行的指令是否出現錯誤；ture為正確；false為錯誤
	static int CurrentErrorDevice = 0;// 統計目前出錯的設備數量
	XSSFSheet Sheet;
	static long totaltime;// 統計所有案例測試時間
	static int location;// 紀錄driver arraylist中null的index
	static int CurrentCase;

	public static void main(String[] args) throws IOException, NoSuchMethodException, SecurityException,
			IllegalAccessException, IllegalArgumentException, InvocationTargetException, InstantiationException {
		initial();
		CreateAppiumSession();// 建立 Appium 通道
		invokeFunction();
		EndAppiumSession();// 中斷 Appium 通道
		System.out.println("測試結束!!!" + "(" + totaltime + " s)");
		Process proc = Runtime.getRuntime().exec("explorer C:\\TUTK_QA_TestTool\\TestReport");// 開啟TestReport資料夾

	}

	public static void initial() {// 初始化CaseErrorList矩陣
		for (int i = 0; i < CaseErrorList.length; i++) {
			for (int j = 0; j < CaseErrorList[i].length; j++) {
				CaseErrorList[i][j] = "";// 填入空字串，避免取值時，出現錯誤
			}
		}
		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			ErrorList[i] = "";
		}
	}

	public static void invokeFunction() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, InstantiationException {
		Object obj = new method();
		Class c = obj.getClass();
		String methodName = null;

		for (CurrentCase = 0; CurrentCase < TestCase.StepList.size(); CurrentCase++) {
			Stopwatch timer = Stopwatch.createStarted();// 開始計時
			System.out.println("[info] CaseName:|" + TestCase.CaseList.get(CurrentCase).toString() + "|");
			CommandError = true;// 預設CommandError為True
			CurrentErrorDevice = 0;// 預設出錯的設備數為0台
			for (int CurrentCaseStep = 0; CurrentCaseStep < TestCase.StepList.get(CurrentCase)
					.size(); CurrentCaseStep++) {
				if (!CommandError & CurrentErrorDevice == TestCase.DeviceInformation.deviceName.size()) {
					break;// 若目前測試案例出現CommandError=false，則跳出目前案例並執行下一個案例
				}
				switch (TestCase.StepList.get(CurrentCase).get(CurrentCaseStep).toString()) {

				case "LaunchAPP":
					methodName = "LaunchAPP";
					break;

				case "Byid_SendKey":
					methodName = "Byid_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "Byid_Click":
					methodName = "Byid_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_Swipe":
					methodName = "Byid_Swipe";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					toElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "ByXpath_SendKey":
					methodName = "ByXpath_SendKey";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "Byid_Clear":
					methodName = "Byid_Clear";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Clear":
					methodName = "ByXpath_Clear";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Click":
					methodName = "ByXpath_Click";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Swipe":
					methodName = "ByXpath_Swipe";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					toElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					CurrentCaseStep = CurrentCaseStep + 2;
					break;

				case "Byid_Wait":
					methodName = "Byid_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_Wait":
					methodName = "ByXpath_Wait";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "HideKeyboard":
					methodName = "HideKeyboard";
					break;

				case "Byid_VerifyText":
					methodName = "Byid_VerifyText";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_VerifyText":
					methodName = "ByXpath_VerifyText";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Sleep":
					methodName = "Sleep";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ScreenShot":
					methodName = "ScreenShot";
					break;

				case "Orientation":
					methodName = "Orientation";
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Swipe":
					methodName = "Swipe";
					startx = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1));
					starty = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2));
					endx = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					endy = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 4));
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 5));
					CurrentCaseStep = CurrentCaseStep + 5;
					break;

				case "ByXpath_Swipe_Vertical":
					methodName = "ByXpath_Swipe_Vertical";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					CurrentCaseStep = CurrentCaseStep + 3;
					break;

				case "ByXpath_Swipe_Horizontal":
					methodName = "ByXpath_Swipe_Horizontal";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					iterative = Integer.valueOf(TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3));
					CurrentCaseStep = CurrentCaseStep + 3;
					break;

				case "ByXpath_Swipe_FindText_Click_Android":
					methodName = "ByXpath_Swipe_FindText_Click_Android";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					scroll = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 2);
					appElemntarray = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 3);
					appInput = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 4);
					appInputXpath = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 5);
					CurrentCaseStep = CurrentCaseStep + 5;
					break;

				case "Byid_LongPress":
					methodName = "Byid_LongPress";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_LongPress":
					methodName = "ByXpath_LongPress";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Byid_invisibility":
					methodName = "Byid_invisibility";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "ByXpath_invisibility":
					methodName = "ByXpath_invisibility";
					appElemnt = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				case "Back":
					methodName = "Back";
					break;

				case "Home":
					methodName = "Home";
					break;

				case "Power":
					methodName = "Power";
					break;

				case "Recent":
					methodName = "Recent";
					break;

				case "ResetAPP":
					methodName = "ResetAPP";
					break;

				case "QuitAPP":
					methodName = "QuitAPP";
					checkVerifyText = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep - 2).toString();
					break;

				case "WiFi":
					methodName = "WiFi";
					switchWiFi = TestCase.StepList.get(CurrentCase).get(CurrentCaseStep + 1);
					CurrentCaseStep = CurrentCaseStep + 1;
					break;

				}

				Method method;
				method = c.getMethod(methodName, new Class[] {});
				method.invoke(c.newInstance());

			}
			System.out.println("[info] Time:|" + timer.stop() + "|");
			totaltime += timer.elapsed(TimeUnit.SECONDS);
			System.out.println("");
			ResetDriverArrayList();// reset driver arraylist
		}
		// System.out.println("測試結束!!!" + "(" + totaltime + " s)");
	}

	public static void ErrorCheck(Object... elements) throws IOException {
		DateFormat df = new SimpleDateFormat("MMM dd, yyyy h:mm:ss a");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);

		// APP畫面上找不到指定元件
		if ((elements.length - 1) > 1) {
			String APPElement = "";
			int i = 0;
			for (Object element : elements) {
				APPElement = APPElement + element;
				if (i != (elements.length - 2)) {// 判斷是否處理到倒數第二個element
					APPElement = APPElement + " or ";// 組成" A元件 or B元件 or
														// C元件"字串
				} else {
					break;// 是倒數第二個break for
				}
				i++;
			}
			System.err.print("[Error] Can't find " + APPElement + " on screen.");
		} else {
			int j = 0;
			for (Object element : elements) {
				if (j != (elements.length - 1)) {// 判斷是否處理到最後一個element

					if (element.equals("On")) {// ok
						System.err.print("[Error] Can't turn on WiFi. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Off")) {// ok
						System.err.print("[Error] Can't turn off WiFi. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("HideKeyboard")) {// ok
						System.out.print("[Error] Can't hide keyboard. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Sleep")) {// ok
						System.err.print("[Error] Fail to sleep. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("ScreenShot")) {// ok
						System.err.print("[Error] Fail to ScreenShot. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Portrait")) {// ok
						System.err.print("[Error] Can't rotate to portrait. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Landscape")) {// ok
						System.err.print("[Error] Can't rotate to landscape. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("EndAppiumSession")) {// ok
						System.err.print("[Error] Can't end session. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("QuitAPP")) {// ok
						System.err.print("[Error] Can't close APP. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("ResetAPP")) {// ok
						System.err.print("[Error] Can't reset APP. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("CeateAppiumSession")) {// ok
						System.out.print("[Error] Can't create new session. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("LaunchAPP")) {// ok
						System.err.print("[Error] Can't launch APP. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("BACK") || element.equals("Home") || element.equals("Power")
							|| element.equals("Recent")) {// ok
						System.err.print("[Error] Can't execute " + element + " button. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Customized")) {// ok
						System.err.print("[Error] Can't execute Customized_Method. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else if (element.equals("Swipe")) {// ok

						System.err.print("[Error] Can't swipe " + "(" + startx + "," + starty + ")" + " to (" + endx
								+ "," + endy + "). (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					} else {
						System.err.print("[Error] Can't find " + element + " on screen. (port:"
								+ (Integer.parseInt(String.valueOf(elements[elements.length - 1])) * 2 + StartPort)
								+ ")");
					}
				}
				j++;
			}
		}
		System.err.println(" " + reportDate);
		String FilePath = MakeErrorFolder(Integer.parseInt(String.valueOf(elements[elements.length - 1])));// 建立各案例之行動裝置資料夾存放log資訊及Screenshot資訊
		logcat(FilePath, Integer.parseInt(String.valueOf(elements[elements.length - 1])));// 收集log
		ErrorScreenShot(FilePath, Integer.parseInt(String.valueOf(elements[elements.length - 1])));// Screenshot
																									// Error畫面
		ErrorList[Integer.parseInt(String.valueOf(elements[elements.length - 1]))] = "Error";// 儲存第i台設備command失敗結果
		CaseErrorList[CurrentCaseNumber] = ErrorList;// 儲存第i台設備執行第CurrentCaseNumber個案例之command失敗結果
		CommandError = false;// 若找不到指定元件，則設定CommandError=false
		driver.set(Integer.parseInt(String.valueOf(elements[elements.length - 1])), null);// 將出錯的driver設定為null
		CurrentErrorDevice++;// 統計出錯的設備數
	}

	public static void logcat(String FilePath, int DeviceNum) throws IOException {
		// 收集log
		// System.out.println(
		// "[info] Executing:|Saving device log...|" +
		// TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
		DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);
		List<LogEntry> logEntries = driver.get(DeviceNum).manage().logs().get("logcat").filter(Level.ALL);
		try {
			FileWriter fw = new FileWriter(FilePath + reportDate + "_log" + ".txt");
			for (int i = 0; i < logEntries.size(); i++) {
				fw.write(logEntries.get(i).toString() + "\n");
			}
			fw.flush();
			fw.close();
			System.out.println("[info] Executing:|Saving device log - Done.|"
					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
		} catch (Exception e) {
			System.err.println("[Error] Executing:|Fail to save device log.|"
					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
		}
	}

	public static void ErrorScreenShot(String FilePath, int DeviceNum) {
		try {
			System.out.println("[info] Executing:|Taking a screenshot of error.|"
					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
			DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
			Date today = Calendar.getInstance().getTime();
			String reportDate = df.format(today);
			File screenShotFile = (File) driver.get(DeviceNum).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(screenShotFile, new File(FilePath + reportDate + ".jpg"));
		} catch (IOException e) {
			System.err.println("[Error] Executing:|Fail to ErrorScreenShot.|"
					+ TestCase.DeviceInformation.deviceName.get(DeviceNum) + "|");
		}
	}

	public static String MakeErrorFolder(int i) {
		// 資料夾結構
		// C:\TUTK_QA_TestTool\TestReport\appPackage\CaseName\DeviceUdid\log\
		String filePath = "C:\\TUTK_QA_TestTool\\TestReport\\" + TestCase.DeviceInformation.appPackage.toString() + "\\"
				+ TestCase.CaseList.get(CurrentCase).toString() + "\\" + TestCase.DeviceInformation.deviceName.get(i)
				+ "\\log\\";
		File file = new File(filePath);
		if (!file.exists()) {
			file.mkdirs();
		}
		return filePath;
	}

	public static void ResetDriverArrayList() {
		// 單一Case結束後，將driver arraylist中null的部分回復為原本driver資料(由driverBK
		// arraylist匯入driver arraylist)
		boolean arraycheck;
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) == null) {
				location = i;// 紀錄driver arraylist中null index位址
				for (int k = 0; k < driverBK.size(); k++) {
					arraycheck = false;
					for (int l = 0; l < driver.size(); l++) {
						// 確認driverBK arraylist資料是否存在於driver arraylist中
						if (driverBK.get(k).equals(driver.get(l))) {
							arraycheck = false;
							break;
						} else {
							arraycheck = true;
						}
					}
					if (arraycheck) {
						driver.set(location, driverBK.get(k));// 覆蓋driver
																// arraylist中的null
					}
				}
			}
		}

	}

	public void Byid_VerifyText() throws IOException {
		boolean result[] = new boolean[driver.size()];// 未給定Boolean值，預設為False
		boolean ErrorResult[] = new boolean[driver.size()];

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_VerifyText|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt))).getText();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}

				if (element[i].equals("ERROR")) {
					ErrorResult[i] = true;

				} else {
					// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
					ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
					for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
						if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
							result[i] = true;
							break;
						} else {
							result[i] = false;
						}
					}
				}
			}
		}
		SubMethod_Result(ErrorResult, result);// 呼叫submethod_result儲存測試結果於Excel

	}

	public void ByXpath_VerifyText() throws IOException {
		boolean result[] = new boolean[driver.size()];// 未給定Boolean值，預設為False
		boolean ErrorResult[] = new boolean[driver.size()];

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_VerifyText|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					element[i] = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)))
							.getText();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
					// System.err.println("[Error] Can't find " + appElemnt);
					// element[i] = "ERROR";// 找不到該物件，回傳Error
					// driver.set(i, null);
					// CurrentErrorDevice++;// 統計設出錯的設備數
				}

				if (element[i].equals("ERROR")) {
					ErrorResult[i] = true;

				} else {
					// 回傳測試案例清單的名稱給ExpectResult.LoadExpectResult，並存放期望結果至ResultList清單
					ExpectResult.LoadExpectResult(TestCase.CaseList.get(CurrentCaseNumber).toString());
					for (int j = 0; j < ExpectResult.ResultList.size(); j++) {
						if (element[i].equals(ExpectResult.ResultList.get(j)) == true) {
							result[i] = true;
							break;
						} else {
							result[i] = false;
						}
					}
				}
			}
		}
		SubMethod_Result(ErrorResult, result);

	}

	public void Byid_Wait() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_Wait|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.presenceOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)));
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_Wait() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Wait|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.presenceOfElementLocated(By.xpath(appElemnt)));
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void Byid_SendKey() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_SendKey|" + appElemnt + "|" + appInput + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					// wait[i].until(ExpectedConditions.visibilityOfElementLocated(
					// By.id(TestCase.DeviceInformation.appPackage + ":id/" +
					// appElemnt))).clear();
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt))).sendKeys(appInput);
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
					HideKeyboard();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_SendKey() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_SendKey|" + appElemnt + "|" + appInput + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					// wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).clear();
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)))
							.sendKeys(appInput);
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
					HideKeyboard();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void Byid_Click() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|Byid_Click|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt))).click();
					ErrorList[i] = "Pass";// 儲存第i台設備command結果，成功執行Click則，存入Pass||舉例
											// 迭代1：ErrorList=[Pass];迭代2：ErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// 儲存第i台設備執行第CurrentCaseNumber個案例之command結果||舉例
																	// 迭代1：CaseErrorList=[[Pass]];迭代2：CaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}

	}

	public void ByXpath_Click() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Click|" + appElemnt + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).click();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void Byid_Clear() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_Clear|" + appElemnt + "|Clear|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt))).clear();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
					HideKeyboard();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}

	}

	public void ByXpath_Clear() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_Clear|" + appElemnt + "|Clear|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))).clear();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
					HideKeyboard();
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}

	}

	public void HideKeyboard() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					if (driver.get(i).isKeyboardShown()) {
						System.out.println(
								"[info] Executing:|HideKeyboard|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
						driver.get(i).hideKeyboard();
					}
				} catch (Exception e) {
					ErrorCheck("HideKeyboard", i);
				}
			}
		}
	}

	public void Sleep() throws IOException {
		// String NewString = "";// 新字串
		// char[] r = { '.' };// 小數點字元
		// char[] c = appInput.toCharArray();// 將字串轉成字元陣列
		// for (int i = 0; i < c.length; i++) {
		// if (c[i] != r[0]) {// 判斷字元是否為小數點
		// NewString = NewString + c[i];// 否，將字元組合成新字串
		// } else {
		// break;// 是，跳出迴圈
		// }
		// }
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {

					System.out.println("[info] Executing:|Sleep|" + appInput + " second..." + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					Thread.sleep((long) (Float.valueOf(appInput) * 1000));
					// Thread.sleep(Integer.valueOf(NewString) * 1000);//
					// 將字串轉成整數

				} catch (Exception e) {
					ErrorCheck("Sleep", i);
				}
			}
		}
	}

	public void ScreenShot() throws IOException {
		DateFormat df = new SimpleDateFormat("yyyy_MM_dd_HH-mm-ss");
		Date today = Calendar.getInstance().getTime();
		String reportDate = df.format(today);

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					File screenShotFile = (File) driver.get(i).getScreenshotAs(OutputType.FILE);
					System.out.println(
							"[info] Executing:|ScreenShot|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					FileUtils.copyFile(screenShotFile, new File("C:\\TUTK_QA_TestTool\\TestReport\\" + "#" + i + 1
							+ TestCase.CaseList.get(CurrentCaseNumber) + "_" + reportDate + ".jpg"));
				} catch (IOException e) {
					ErrorCheck("ScreenShot", i);
				}
			}
		}
	}

	public void Orientation() throws IOException {

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					if (appInput.equals("Landscape")) {
						System.out.println("[info] Executing:|Orientation|Landscape|"
								+ TestCase.DeviceInformation.deviceName.get(i) + "|");
						driver.get(i).rotate(ScreenOrientation.LANDSCAPE);
					} else if (appInput.equals("Portrait")) {
						System.out.println("[info] Executing:|Orientation|Portrait|");
						driver.get(i).rotate(ScreenOrientation.PORTRAIT);
					}
				} catch (Exception e) {
					if (appInput.equals("Landscape")) {
						ErrorCheck("Landscape", i);
					} else {
						ErrorCheck("Portrait", i);
					}
				}
			}
		}
	}

	public static void EndAppiumSession() throws IOException {
		for (int j = 0; j < driver.size(); j++) {
			if (driver.get(j) != null) {
				try {
					System.out.println("[info] Executing:|End Session|Server Port:" + (port - 2 - (j * 2)) + "|");
					driver.get(j).quit();
				} catch (Exception e) {
					ErrorCheck("EndAppiumSession", j);
				}
			}
		}
	}

	public void QuitAPP() throws IOException {
		boolean state[] = new boolean[driver.size()];
		for (int j = 0; j < driver.size(); j++) {
			if (driver.get(j) != null) {
				try {
					System.out
							.println("[info] Executing:|QuitAPP|" + TestCase.DeviceInformation.deviceName.get(j) + "|");
					driver.get(j).closeApp();

					// 執行完QuitAPP後，代表該Case流程操作正常結束，因此會再TestReport填入Pass
					// ;但若先前執行過VerifyText動作後，也會在TestReport填入該Case結果Pass/Fail/Error，因此此處需判斷是否執行過Verify，避免Verify結果被刷掉，如下判斷式
					for (int i = 0; i < TestCase.StepList.get(CurrentCaseNumber).size(); i++) {
						if (TestCase.StepList.get(CurrentCaseNumber).get(i).equals("Byid_VerifyText")
								|| TestCase.StepList.get(CurrentCaseNumber).get(i).equals("ByXpath_VerifyText")
								|| TestCase.StepList.get(CurrentCaseNumber).get(i).equals("Byid_VerifyRadioButton")
								|| TestCase.StepList.get(CurrentCaseNumber).get(i)
										.equals("ByXpath_VerifyRadioButton")) {
							state[j] = true;// true代表找到Verify
							break;
						}
					}
					if (!state[j]) {// 若state[j]=false表該case未執行Verify指令，故執行以下流程

						try {
							workBook = new XSSFWorkbook(
									new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
						} catch (Exception e) {
							System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
						}

						if (TestCase.DeviceInformation.deviceName.get(j).toString().length() > 20) {// Excel工作表名稱最常31字元因，故需判斷UDID長度是否大於31
							char[] NewUdid = new char[20];// 因需包含_TestReport字串(共11字元)，故設定20位字元陣列(31-11)
							TestCase.DeviceInformation.deviceName.get(j).toString().getChars(0, 20, NewUdid, 0);// 取出UDID前20字元給NewUdid
							Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// 根據NewUdid，指定某台裝置的TestReport
																								// sheet
						} else {
							Sheet = workBook
									.getSheet(TestCase.DeviceInformation.deviceName.get(j).toString() + "_TestReport");// 指定某台裝置的TestReport
																														// sheet
						}
						if (CaseErrorList[CurrentCaseNumber][j].equals("Pass")) {// 取出CaseErrorList之第CurrentCaseNumber個測項中的第i台行動裝置之結果
							Sheet.getRow(CurrentCaseNumber + 1).getCell(1).setCellValue("Pass");// 填入第i台行動裝置之第CurrentCaseNumber個測項結果Pass
						}
						// 執行寫入Excel後的存檔動作
						try {
							FileOutputStream out = new FileOutputStream(
									new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
							workBook.write(out);
							out.close();
							workBook.close();
						} catch (Exception e) {
							System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
						}
					}

				} catch (Exception e) {
					ErrorCheck("QuitAPP", j);
				}
			}
		}
	}

	public void ResetAPP() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println(
							"[info] Executing:|ResetAPP|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).resetApp();
				} catch (Exception e) {
					ErrorCheck("ResetAPP", i);
				}
			}
		}
	}

	public static void CreateAppiumSession() throws IOException {
		DesiredCapabilities cap[] = new DesiredCapabilities[TestCase.DeviceInformation.deviceName.size()];
		driver.clear();// 清空driver arraylist
		
		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			cap[i] = new DesiredCapabilities();
		}

		for (int i = 0; i < TestCase.DeviceInformation.deviceName.size(); i++) {
			System.out.println("DEVICE_NAME:" + TestCase.DeviceInformation.deviceName.get(i).toString());
			cap[i].setCapability(MobileCapabilityType.NEW_COMMAND_TIMEOUT, device_timeout);
			cap[i].setCapability(MobileCapabilityType.DEVICE_NAME, TestCase.DeviceInformation.deviceName.get(i));
			cap[i].setCapability(MobileCapabilityType.UDID, TestCase.DeviceInformation.deviceName.get(i));
			cap[i].setCapability(MobileCapabilityType.PLATFORM_VERSION,
					TestCase.DeviceInformation.platformVersion.get(i));
			cap[i].setCapability(AndroidMobileCapabilityType.APP_PACKAGE, TestCase.DeviceInformation.appPackage);
			cap[i].setCapability(AndroidMobileCapabilityType.APP_ACTIVITY, TestCase.DeviceInformation.appActivity);
			cap[i].setCapability(MobileCapabilityType.NO_RESET, TestCase.DeviceInformation.ResetAPP);// 每次單一Case執行前，是否清除APP快取資料;是為true;否為false
			cap[i].setCapability("autoLaunch", false);
			if (TestCase.DeviceInformation.UIAutomator2) {
				cap[i].setCapability(MobileCapabilityType.AUTOMATION_NAME, "uiautomator2");// 因Android
																							// 7.0設備以上需使用UIAutomator2，才可操作UI元件
			}
			// cap[i].setCapability(AndroidMobileCapabilityType.SYSTEM_PORT,
			// systemPort);

			try {

				System.out.println("[info] Executing:|Create New Session|Server Port:" + port + "|");
				Androiddriver = new AndroidDriver<>(new URL("http://127.0.0.1:" + port + "/wd/hub"), cap[i]);
				driverBK.add(Androiddriver);
				driver.add(Androiddriver);
			} catch (Exception e) {
				ErrorCheck("CeateAppiumSession", i);
			}
			port = port + 2;
			// systemPort = systemPort + 4;
		}
		System.out.println("");
	}

	public void LaunchAPP() throws IOException {
		CurrentCaseNumber = CurrentCaseNumber + 1;
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|LaunchAPP|" + TestCase.DeviceInformation.appPackage + "|"
							+ TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).launchApp();
				} catch (Exception e) {
					ErrorCheck("LaunchAPP", i);
				}
			}
		}

	}

	public void Back() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Back|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).pressKeyCode(AndroidKeyCode.BACK);
				} catch (Exception e) {
					ErrorCheck("Back", i);
				}
			}
		}
	}

	public void Home() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Home|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).pressKeyCode(AndroidKeyCode.HOME);
				} catch (Exception e) {
					ErrorCheck("Home", i);
				}
			}
		}
	}

	public void Power() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Power|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).pressKeyCode(AndroidKeyCode.KEYCODE_POWER);
				} catch (Exception e) {
					ErrorCheck("Power", i);
				}
			}
		}
	}

	public void Recent() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out
							.println("[info] Executing:|Recent|" + TestCase.DeviceInformation.deviceName.get(i) + "|");
					driver.get(i).pressKeyCode(AndroidKeyCode.KEYCODE_APP_SWITCH);
				} catch (Exception e) {
					ErrorCheck("Recent", i);
				}
			}
		}
	}

	// public void WiFi() throws IOException {
	//
	// for (int i = 0; i < driver.size(); i++) {
	// if (driver.get(i) != null) {

	// try {
	// System.out.println("[info] Executing:|WiFi|" + switchWiFi + "|"
	// + TestCase.DeviceInformation.deviceName.get(i) + "|");
	// //
	// if邏輯說明:(目的避免已開啟wifi或已關閉wif了，又再次執行令啟動wifi或關閉wif(皆去除if判斷僅跑switch流程)，如此可節省測試時間)
	// // [判斷手機連線狀態為WiFi off及data off(皆NONE) &&
	// // Excel指令為On時，才執行switch之Case"On"]
	// // ||[[判斷手機連線狀態為WiFi on || WiFi及Data都啟動(皆ALL)] &&
	// // Excel指令為Off時，才執行switch之Case"Off"]
	//
	// if ((driver.get(i).getConnection().toString().equals("NONE") &&
	// switchWiFi.equals("On"))
	// || ((driver.get(i).getConnection().toString().equals("ALL")
	// || driver.get(i).getConnection().toString().equals("WIFI"))
	// && switchWiFi.equals("Off"))) {
	// switch (switchWiFi) {
	// case "On":
	//
	// // Appium 6.1.0 API 有問題!!!!
	// driver.get(i).setConnection(Connection.WIFI);
	// break;
	// case "Off":
	// // Appium 6.1.0 API 有問題!!!!
	// driver.get(i).setConnection(Connection.NONE);
	// break;
	// }
	// }
	// } catch (Exception e) {
	// switch (switchWiFi) {
	// case "On":
	// ErrorCheck("On", i);
	// break;
	// case "Off":
	// ErrorCheck("Off", i);
	// break;
	// }
	// }
	// }
	// }
	//
	// }

	public void Byid_invisibility() throws IOException {
		for (int i = 0; i < driver.size(); i++) {

			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|Byid_invisibility|" + appElemnt + "|");

					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.invisibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)));
					ErrorList[i] = "Pass";// 儲存第i台設備command結果，成功執行Click則，存入Pass||舉例
											// 迭代1：ErrorList=[Pass];迭代2：ErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// 儲存第i台設備執行第CurrentCaseNumber個案例之command結果||舉例
																	// 迭代1：CaseErrorList=[[Pass]];迭代2：CaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {

					ErrorCheck(appElemnt, i);

				}
			}
		}
	}

	public void ByXpath_invisibility() throws IOException {
		for (int i = 0; i < driver.size(); i++) {

			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|ByXpath_invisibility|" + appElemnt + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					wait[i].until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(appElemnt)));
					ErrorList[i] = "Pass";// 儲存第i台設備command結果，成功執行Click則，存入Pass||舉例
											// 迭代1：ErrorList=[Pass];迭代2：ErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// 儲存第i台設備執行第CurrentCaseNumber個案例之command結果||舉例
																	// 迭代1：CaseErrorList=[[Pass]];迭代2：CaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void Byid_LongPress() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|Byid_LongPress|" + appElemnt + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);

					t.longPress(LongPressOptions.longPressOptions().withElement(
							ElementOption.element(wait[i].until(ExpectedConditions.visibilityOfElementLocated(
									By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt))))))
							.perform();

					ErrorList[i] = "Pass";// 儲存第i台設備command結果，成功執行Click則，存入Pass||舉例
											// 迭代1：ErrorList=[Pass];迭代2：ErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// 儲存第i台設備執行第CurrentCaseNumber個案例之command結果||舉例
																	// 迭代1：CaseErrorList=[[Pass]];迭代2：CaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}

		}

	}

	public void ByXpath_LongPress() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {

				try {
					System.out.println("[info] Executing:|ByXpath_LongPress|" + appElemnt + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					t.longPress(LongPressOptions.longPressOptions()
							.withElement(ElementOption.element(
									wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt))))))
							.perform();

					ErrorList[i] = "Pass";// 儲存第i台設備command結果，成功執行Click則，存入Pass||舉例
											// 迭代1：ErrorList=[Pass];迭代2：ErrorList=[Pass,Pass]
					CaseErrorList[CurrentCaseNumber] = ErrorList;// 儲存第i台設備執行第CurrentCaseNumber個案例之command結果||舉例
																	// 迭代1：CaseErrorList=[[Pass]];迭代2：CaseErrorList=[[Pass,Pass]]
				} catch (Exception e) {
					ErrorCheck(appElemnt, i);
				}
			}

		}

	}

	public void ByXpath_Swipe() throws IOException {
		Point p1, p2;// p1 為起點;p2為終點

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Swipe|" + appElemnt + "|" + toElemnt + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					WebElement ele2 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(toElemnt)));
					WebElement ele1 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));

					t.press(PointOption.point(ele1.getLocation().x, ele1.getLocation().y))
							.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
							.moveTo(PointOption.point(ele2.getLocation().x, ele2.getLocation().y)).release().perform();
					// t.press(ele1).waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(ele2).release().perform();
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(toElemnt, appElemnt, i);
				}
			}
		}
	}

	public void Byid_Swipe() throws IOException {
		Point p1, p2;// p1 為起點;p2為終點

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|Byid_Swipe|" + appElemnt + "|" + toElemnt + "|");

					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					TouchAction t = new TouchAction(driver.get(i));
					WebElement ele2 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + toElemnt)));
					WebElement ele1 = wait[i].until(ExpectedConditions.visibilityOfElementLocated(
							By.id(TestCase.DeviceInformation.appPackage + ":id/" + appElemnt)));
					t.press(PointOption.point(ele1.getLocation().x, ele1.getLocation().y))
							.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
							.moveTo(PointOption.point(ele2.getLocation().x, ele2.getLocation().y)).release().perform();

					// t.press(ele1).waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(ele2).release().perform();

					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception e) {
					ErrorCheck(toElemnt, appElemnt, i);
				}
			}
		}
	}

	public void Swipe() throws IOException {
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				for (int j = 0; j < iterative; j++) {
					try {
						System.out.println(
								"[info] Executing:|Swipe|(" + startx + "," + starty + ")|(" + endx + "," + endy + ")|");
						TouchAction t = new TouchAction(driver.get(i));

						t.press(PointOption.point(startx, starty)).waitAction(WaitOptions.waitOptions(ofSeconds(1)))
								.moveTo(PointOption.point(endx, endy)).release().perform();
					} catch (Exception e) {
						ErrorCheck("Swipe", i);
						break;// 出錯後，離開iterative回圈
					}
				}
			}
		}
	}

	public void ByXpath_Swipe_Vertical() throws IOException {
		Point p;// 元件座標
		Dimension s;// 元件大小
		WebElement e;

		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Swipe_Vertical|" + appElemnt + "|" + scroll + "|"
							+ iterative + "|");
					TouchAction t = new TouchAction(driver.get(i));
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					e = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));
					s = e.getSize();
					p = e.getLocation();
					int errorX = (int) Math.round(s.width * 0.01);
					int errorY = (int) Math.round(s.height * 0.01);
					for (int j = 0; j < iterative; j++) {
						if (scroll.equals("DOWN")) {// 畫面向下捲動

							t.press(PointOption.point(p.x + errorX, p.y + s.height - errorY))
									.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
									.moveTo(PointOption.point(p.x + errorX, p.y + errorY)).release().perform();

						} else if (scroll.equals("UP")) {// 畫面向上捲動
							t.press(PointOption.point(p.x + errorX, p.y + errorY))
									.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
									.moveTo(PointOption.point(p.x + errorX, p.y + s.height - errorY)).release()
									.perform();
						}
					}
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception w) {
					ErrorCheck(appElemnt, i);
				}
			}
		}
	}

	public void ByXpath_Swipe_Horizontal() throws IOException {
		Point p;// 元件座標
		Dimension s;// 元件大小
		WebElement e;
		for (int i = 0; i < driver.size(); i++) {
			if (driver.get(i) != null) {
				try {
					System.out.println("[info] Executing:|ByXpath_Swipe_Horizontal|" + appElemnt + "|" + scroll + "|"
							+ iterative + "|");
					wait[i] = new WebDriverWait(driver.get(i), command_timeout);
					TouchAction t = new TouchAction(driver.get(i));
					e = wait[i].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));

					s = e.getSize();
					p = e.getLocation();
					int errorX = (int) Math.round(s.getWidth() * 0.01);
					int errorY = (int) Math.round(s.getHeight() * 0.01);
					for (int j = 0; j < iterative; j++) {
						if (scroll.equals("RIGHT")) {// 畫面向右捲動 (觀看畫面左方內容)

							t.press(PointOption.point(p.x + errorX, p.y + errorY))
									.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
									.moveTo(PointOption.point(p.x + s.width - errorX, p.y + errorY)).release()
									.perform();
						} else if (scroll.equals("LEFT")) {// 畫面向左捲動 (觀看畫面右方內容)

							t.press(PointOption.point(p.x + s.width - errorX, p.y + errorY))
									.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
									.moveTo(PointOption.point(p.x + errorX, p.y + errorY)).release().perform();
						}
					}
					ErrorList[i] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception w) {
					ErrorCheck(appElemnt, i);
				}
			}

		}
	}

	// ByXpath_Swipe_FindText_Click_Android缺點：1.搜尋的字串不可重複 2.搜尋5次都沒找到元件，則停止搜尋
	public void ByXpath_Swipe_FindText_Click_Android() throws IOException {

		for (int j = 0; j < driver.size(); j++) {
			if (driver.get(j) != null) {
				int SearchNumber = 0;// 搜尋次數
				Point ScrollBarP;// 卷軸元件座標
				Dimension ScrollBarS;// 卷軸元件之長及寬
				WebElement ScrollBar;// 卷軸元件

				try {
					System.out.println("[info] Executing:|ByXpath_Swipe_FindText_Click_Android|" + appElemnt + "|"
							+ scroll + "|" + appElemntarray + "|" + appInput + "|" + appInputXpath + "|");
					wait[j] = new WebDriverWait(driver.get(j), command_timeout);
					TouchAction t = new TouchAction(driver.get(j));
					ScrollBar = wait[j].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appElemnt)));// 卷軸元件
					ScrollBarS = ScrollBar.getSize();// 卷軸元件的長及寬
					ScrollBarP = ScrollBar.getLocation();// 卷軸的座標
					int errorX = (int) Math.round(ScrollBarS.width * 0.1);
					int errorY = (int) Math.round(ScrollBarS.height * 0.1);
					List<WebElement> targetlist = wait[j]
							.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(appElemntarray)));// 要搜尋的多筆類似元件清單

					for (int i = 0; i < targetlist.size(); i++) {

						if ((targetlist.get(i).getText().toString()).equals(appInput)) {// 若targetelement在targetlist清單中，則點擊targetelement
							WebElement targetElement;// 準備搜尋的元件
							Point targetElementP;// 準備搜尋的元件之座標
							Dimension targetElementS;// 準備搜尋的元件之長及寬

							targetElement = wait[j]
									.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appInputXpath)));

							targetElementP = targetElement.getLocation();// 準備搜尋元件的座標
							targetElementS = targetElement.getSize();// 準備搜尋元件的長及寬

							switch (scroll.toString()) {

							case "DOWN":
								if (targetElementP.y > ScrollBarP.y + ScrollBarS.height) {// 若搜尋元件的y座標大於卷軸範圍，表示搜尋元件全部UI被卷軸遮住

									t.press(PointOption.point(targetElementP.x,
											ScrollBarS.height + ScrollBarP.y - errorY))
											.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
											.moveTo(PointOption.point(targetElementP.x, ScrollBarP.y + errorY))
											.release().perform();

								} else if (targetElementP.y + targetElementS.height == ScrollBarP.y
										+ ScrollBarS.height) {// 若搜尋元件的y座標與寬度總和等於卷軸長度，表示搜尋元件的部分UI被卷軸遮住
									t.press(PointOption.point(targetElementP.x - errorY, targetElementP.y))
											.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
											.moveTo(PointOption.point(targetElementP.x, ScrollBarP.y + errorY))
											.release().perform();

								}
								break;

							case "UP":
								if (targetElementP.y + targetElementS.height < ScrollBarP.y) {// 若搜尋元件的最大y座標小於卷軸y座標，表示搜尋元件全部UI被卷軸遮住
									t.press(PointOption.point(targetElementP.x, ScrollBarP.y + errorY))
											.waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(PointOption
													.point(targetElementP.x, ScrollBarS.height + ScrollBarP.y - errorY))
											.release().perform();

								} else {// 反之，若搜尋元件的最大y座標大於卷軸y座標，表示搜尋元件全部UI被卷軸遮住
									t.press(PointOption.point(targetElementP.x, ScrollBarP.y + errorY))
											.waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(PointOption
													.point(targetElementP.x, ScrollBarP.y + ScrollBarS.height - errorY))
											.release().perform();
								}
								break;

							case "LEFT":// 畫面向左捲動(觀看畫面右方內容)
								if (targetElementP.x > ScrollBarP.x + ScrollBarS.width) {// 若搜尋元件的x座標大於卷軸範圍，表示搜尋元件全部UI被卷軸遮住
									t.press(PointOption.point(ScrollBarP.x + ScrollBarS.width - errorX,
											targetElementP.y)).waitAction(WaitOptions.waitOptions(ofSeconds(1)))
											.moveTo(PointOption.point(ScrollBarP.x + errorX, targetElementP.y))
											.release().perform();

								} else if (targetElementP.x + targetElementS.width == ScrollBarP.x + ScrollBarS.width) {// 若搜尋元件的x座標與寬度總和等於卷軸寬度，表示搜尋元件的部分UI被卷軸遮住
									t.press(PointOption.point(targetElementP.x - errorX, targetElementP.y))
											.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
											.moveTo(PointOption.point(ScrollBarP.x + errorX, targetElementP.y))
											.release().perform();
								}
								break;

							case "RIGHT":// 畫面向右捲動(觀看畫面左方內容)
								if (targetElementP.x + targetElementS.width < ScrollBarP.x) {// 若搜尋元件的最大x座標小於卷軸x座標，表示搜尋元件全部UI被卷軸遮住

									t.press(PointOption.point(ScrollBarP.x + errorX, targetElementP.y))
											.waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(PointOption
													.point(ScrollBarP.x + ScrollBarS.width - errorX, targetElementP.y))
											.release().perform();
								} else if (targetElementP.x == ScrollBarP.x) {// 若搜尋元件的x座標等於卷軸x座標，可能表示搜尋元件的部分UI被卷軸遮住
									t.press(PointOption.point(targetElementP.x + targetElementS.width + errorX,
											targetElementP.y)).waitAction(WaitOptions.waitOptions(ofSeconds(1)))
											.moveTo(PointOption.point(ScrollBarP.x + ScrollBarS.width - errorX,
													targetElementP.y))
											.release().perform();
								}
								break;
							}

							wait[j].until(ExpectedConditions.visibilityOfElementLocated(By.xpath(appInputXpath)))
									.click();
							break;
						}

						if (i == targetlist.size() - 1) {// 若targetlist中最後一筆資料比對完後，則進行Srcoll拖曳

							switch (scroll.toString()) {

							case "DOWN":
								t.press(PointOption.point(ScrollBarP.x + errorX,
										ScrollBarP.y + ScrollBarS.height - errorY))
										.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
										.moveTo(PointOption.point(ScrollBarP.x + errorX, ScrollBarP.y + errorY))
										.release().perform();// 向下捲動
								break;

							case "UP":

								t.press(PointOption.point(ScrollBarP.x + errorX, ScrollBarP.y + errorY))
										.waitAction(WaitOptions.waitOptions(ofSeconds(1)))
										.moveTo(PointOption.point(ScrollBarP.x + errorX,
												ScrollBarP.y + ScrollBarS.height - errorY))
										.release().perform();// 向上捲動
								break;

							case "LEFT":
								t.press(PointOption.point(ScrollBarP.x + ScrollBarS.width - errorX,
										ScrollBarP.y + errorY)).waitAction(WaitOptions.waitOptions(ofSeconds(1)))
										.moveTo(PointOption.point(ScrollBarP.x + errorX, ScrollBarP.y + errorY))
										.release().perform();// 畫面向左捲動(觀看畫面右方內容)
								break;

							case "RIGHT":
								t.press(PointOption.point(ScrollBarP.x + errorX, ScrollBarP.y + errorY))
										.waitAction(WaitOptions.waitOptions(ofSeconds(1))).moveTo(PointOption
												.point(ScrollBarP.x + ScrollBarS.width - errorX, ScrollBarP.y + errorY))
										.release().perform();// 畫面向右捲動(觀看畫面左方內容)
								break;

							}
							SearchNumber++;// 累計搜尋次數
							targetlist.clear();// 清除targetlist
							targetlist = wait[j].until(
									ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(appElemntarray)));// 再次取得新targetlist

							if (SearchNumber == 10) {// 搜尋10次都沒找到元件，則跳出for
								System.err.println("Can't find " + appInput);// 印出找不到
								break;// 跳出for
							} else {
								i = -1;// 若SearchNumber!=10，則令i=-1(目的：再次執行for)
							}
						}
					}
					ErrorList[j] = "Pass";
					CaseErrorList[CurrentCaseNumber] = ErrorList;
				} catch (Exception w) {

					ErrorCheck(appElemnt, appElemntarray, appInputXpath, j);

				}
			}
		}

	}

	/*
	 * 上下隨機滑動n次 public void Swipe() { Random rand = new Random(); boolean
	 * items[] = { true, false }; for (int i = 0; i < driver.size(); i++) { for
	 * (int j = 0; j < iterative; j++) { if (items[rand.nextInt(items.size())])
	 * { driver.get(i).swipe(startx, starty, endx, endy, 500); }else{
	 * driver.get(i).swipe(endx, endy, startx , starty , 500); } } } }
	 * 
	 */

	public void SubMethod_Result(boolean ErrorResult[], boolean result[]) {
		// 開啟Excel
		try {
			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
		} catch (Exception e) {
			System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
		}
		for (int i = 0; i < driver.size(); i++) {

			if (TestCase.DeviceInformation.deviceName.get(i).toString().length() > 20) {// Excel工作表名稱最常31字元因，故需判斷UDID長度是否大於31
				char[] NewUdid = new char[20];// 因需包含_TestReport字串(共11字元)，故設定20位字元陣列(31-11)
				TestCase.DeviceInformation.deviceName.get(i).toString().getChars(0, 20, NewUdid, 0);// 取出UDID前20字元給NewUdid
				Sheet = workBook.getSheet(String.valueOf(NewUdid) + "_TestReport");// 根據NewUdid，指定某台裝置的TestReport
																					// sheet
			} else {
				Sheet = workBook.getSheet(TestCase.DeviceInformation.deviceName.get(i).toString() + "_TestReport");// 指定某台裝置的TestReport
																													// sheet
			}

			if (ErrorResult[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Error");
			} else if (result[i] == true) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Pass");
			} else if (result[i] == false) {
				Sheet.getRow(CurrentCaseNumber + 1).createCell(1).setCellValue("Fail");
			}
		}
		// 執行寫入Excel後的存檔動作
		try {
			FileOutputStream out = new FileOutputStream(new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));
			workBook.write(out);
			out.close();
			workBook.close();
		} catch (Exception e) {
			System.err.println("[Error] Can't find C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm");
		}
	}
}
