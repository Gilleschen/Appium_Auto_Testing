package Appium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoadTestCase {

	public ArrayList<ArrayList<String>> StepList = new ArrayList<ArrayList<String>>();// 所有測試案例的動作清單(2維陣列)
	ArrayList<String> StepListData = new ArrayList<String>();// 單一測試案例的動作清單
	ArrayList<String> CheckStepListData = new ArrayList<String>();// 去除StepListData內空白資料，更新單一測試案例的動作清單
	public ArrayList<String> CaseList = new ArrayList<String>();// 所有測試案例的名稱清單
	LoadDeviceInformation DeviceInformation = new LoadDeviceInformation();

	public void SaveAllCase(XSSFSheet sheet) {
		int i = 0;
		try {

			do {// column Number

				for (int j = 0; j < sheet.getRow(i).getPhysicalNumberOfCells(); j++) {

					if (sheet.getRow(i).getCell(j) != null) {// Apache
																// POI
																// 讀取Excel儲存格時，有機率將空白儲存格讀入，因此需判斷儲存格是否為空白，皆null

						if (sheet.getRow(i).getCell(j).toString().equals("CaseName")) {
							CaseList.add(sheet.getRow(i).getCell(1).toString());// 從指定待測試腳本的sheet中儲存測試案例的名稱
							break;
						} else {

							StepListData.add(sheet.getRow(i).getCell(j).getStringCellValue());// 從指定待測試腳本的sheet中儲存測試案例的步驟
																								// Excel數字要轉成字串型態
						}

					}
				}
				// 判斷單一測試案例是否結束，若是，則StepListData加入StepList
				if (sheet.getRow(i).getCell(0).toString().equals("QuitAPP")
						|| sheet.getRow(i).getCell(0).toString().equals("Quit")) {

					for (int j = 0; j < StepListData.size(); j++) {
						if (StepListData.get(j).toString() != "") {
							CheckStepListData.add(StepListData.get(j).toString());
						}
					}

					StepList.add(CheckStepListData);
					StepListData = new ArrayList<String>();
					CheckStepListData = new ArrayList<String>();
				}
				i++;
			} while (!sheet.getRow(i).getCell(0).toString().equals(""));
		} catch (Exception e) {
			;
		}
	}

	public void SaveCase(XSSFSheet APPDeviceSheet, XSSFSheet sheet, int k) {
		String casename = APPDeviceSheet.getRow(k + 1).getCell(5).toString();
		casename = casename.replaceAll("\\s+", "");// 移除字串內所有空格
		String[] TargetCaseList = casename.split(",");// 將測試案例加入TargetCaseList矩陣

		for (int i = 0; i < TargetCaseList.length; i++) {
			CaseList.add(TargetCaseList[i]);
		}

		for (int c = 0; c < TargetCaseList.length; c++) {
			int i = 0;
			try {
				do {

					if (sheet.getRow(i).getCell(0).toString().equals("CaseName")) {

						if (sheet.getRow(i).getCell(1).toString().equals(TargetCaseList[c].toString())) {

							int w = i + 1;
							try {
								do {
									for (int j = 0; j < sheet.getRow(w).getPhysicalNumberOfCells(); j++) {
										StepListData.add(sheet.getRow(w).getCell(j).getStringCellValue());
									}
									w++;
								} while (!sheet.getRow(w).getCell(0).toString().equals("CaseName"));
								i = w - 1;
							} catch (Exception e) {
								i = w - 1;
							}
							// 判斷單一測試案例是否結束，若是，則StepListData加入StepList
							if (sheet.getRow(i).getCell(0).toString().equals("QuitAPP")
									|| sheet.getRow(i).getCell(0).toString().equals("Quit")) {

								for (int j = 0; j < StepListData.size(); j++) {
									if (StepListData.get(j).toString() != "") {
										CheckStepListData.add(StepListData.get(j).toString());
									}
								}

								StepList.add(CheckStepListData);
								StepListData = new ArrayList<String>();
								CheckStepListData = new ArrayList<String>();
								break;

							}
						}
					}
					i++;
				} while (!sheet.getRow(i).getCell(0).toString().equals(""));

			} catch (Exception e) {
				;
			}

		}

	}

	public LoadTestCase() {
		XSSFWorkbook workbook = null;
		XSSFSheet sheet, APPDeviceSheet;

		try {
			workbook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestTool\\TestScript.xlsm"));

			CaseList = new ArrayList<String>();
			StepList = new ArrayList<ArrayList<String>>();
			StepListData = new ArrayList<String>();

			for (int k = 0; k < DeviceInformation.ScriptList.size(); k++) {

				sheet = workbook.getSheet(DeviceInformation.ScriptList.get(k).toString());// 指定待測試腳本的sheet
				APPDeviceSheet = workbook.getSheet("APP&Device");// hard code
				// System.out.println(APPDeviceSheet.getRow(k +
				// 1).getCell(5).toString());
				if (APPDeviceSheet.getRow(k + 1).getCell(5) == null) {
					SaveAllCase(sheet);
				} else {
					SaveCase(APPDeviceSheet, sheet, k);
				}
			}
		} catch (Exception e) {
			;
		}

		System.out.println("測試腳本：" + StepList);
		System.out.println("");
		// 建立各裝置的Test Report
		for (int i = 0; i < DeviceInformation.deviceName.size(); i++) {

			if (DeviceInformation.deviceName.get(i).toString().length() > 20) {// Excel工作表名稱最常31字元因，故需判斷UDID長度是否大於31
				char[] NewUdid = new char[20];// 因需包含_TestReport字串(共11字元)，故設定20位字元陣列(31-11)
				DeviceInformation.deviceName.get(i).toString().getChars(0, 20, NewUdid, 0);// 取出UDID前20字元給NewUdid
				sheet = workbook.createSheet(String.valueOf(NewUdid) + "_TestReport");// 使用NewUdid命名新工作表
			} else {
				sheet = workbook.createSheet(DeviceInformation.deviceName.get(i).toString() + "_TestReport");
			}

			sheet.createRow(0).createCell(0).setCellValue("CaseName");
			sheet.getRow(0).createCell(1).setCellValue("Result");
			for (int k = 0; k < CaseList.size(); k++) {
				sheet.createRow(k + 1).createCell(0).setCellValue(CaseList.get(k).toString());// 填入各案例名稱
				sheet.getRow(k + 1).createCell(1).setCellValue("Error");// 預設各案列結果為Error

			}
		}
		// 執行寫入Excel後的存檔動作
		FileOutputStream out;
		try {
			out = new FileOutputStream(new File("C:\\TUTK_QA_TestTool\\TestReport\\TestReport.xlsm"));// 另存新檔
			workbook.write(out);
			out.close();
			workbook.close();
		} catch (Exception e) {
			;
		}

	}

	// public static void main(String[] args) {
	// LoadTestCase d = new LoadTestCase();
	// }
}
