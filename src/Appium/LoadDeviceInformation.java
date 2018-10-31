package Appium;

import java.io.FileInputStream;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoadDeviceInformation {

	public ArrayList<String> deviceName = new ArrayList<String>();// 存放手機UDID
	public ArrayList<String> platformVersion = new ArrayList<String>();// 存放手機OS版本
	public ArrayList<String> ScriptList = new ArrayList<String>();// 存放待執行的腳本名稱
	public ArrayList<String> CasetList = new ArrayList<String>();// 存放指定的待側項
	public String appPackage;// APP Packagename
	public String appActivity;// APP Activity
	public Boolean ResetAPP; //每次單一Case執行前，是否清除APP快取資料;是為true;否為false
	public Boolean UIAutomator2; //是否調用UIautomator2

	public LoadDeviceInformation() {
		XSSFWorkbook workBook;
		XSSFSheet Sheet;

		try {

			workBook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestTool\\TestScript.xlsm"));// hard
			// code
			Sheet = workBook.getSheet("APP&Device");// hard code
			deviceName = new ArrayList<String>();
			platformVersion = new ArrayList<String>();
			ScriptList = new ArrayList<String>();
			CasetList = new ArrayList<String>();

			appPackage = Sheet.getRow(1).getCell(0).toString();// hard code
			appActivity = Sheet.getRow(1).getCell(1).toString();// hard code

			if (Sheet.getRow(1).getCell(7).toString() == "TRUE") {
				ResetAPP = false;
			} else {
				ResetAPP = true;
			}
			
			if (Sheet.getRow(1).getCell(8).toString() == "TRUE") {
				UIAutomator2 = true;
			} else {
				UIAutomator2 = false;
			}

			int j = 1;
			try {
				do {
					deviceName.add(Sheet.getRow(j).getCell(2).toString());// 手機UDID
					platformVersion.add(Sheet.getRow(j).getCell(3).toString());// 手機OS版本
					j++;
				} while (!Sheet.getRow(j).getCell(2).toString().equals(""));
			} catch (Exception e) {
				;
			}

			int k = 1;
			try {
				do {
					ScriptList.add(Sheet.getRow(k).getCell(4).toString());// 準備要Run的腳本
					k++;
				} while (!Sheet.getRow(k).getCell(4).toString().equals(""));
			} catch (Exception e) {
				;
			}

			System.out.println("測試設備(UDID/Android Version)");

			for (int i = 0; i < deviceName.size(); i++) {
				System.out.print(deviceName.get(i) + "/");
				System.out.println(platformVersion.get(i));
			}
			workBook.close();

		} catch (Exception e) {

		}
	}

}
