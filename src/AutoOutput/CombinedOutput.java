package AutoOutput;

import com.agile.api.*;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import record.LogNew;
import util.Ini;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class CombinedOutput {
	public static void start(IChange change, IAgileSession admin, Ini ini, LogNew log, String str_ExcelTarget,
			String str_returnTarget) throws Exception {
		try {
			change = (IChange) admin.getObject(IChange.OBJECT_TYPE, change.getName());
			log.log(1, "Get Change as Admin:" + change.getName());
			ITable itable_affectedTable = AutoGetInf.getAffectedTable(change);
			Iterator<?> it = itable_affectedTable.iterator();
			IRow row;
			IItem item;

			// Create Item Rule Map
			ArrayList<IItem> itemlist = new ArrayList<>();
			while (it.hasNext()) {
				row = (IRow) it.next();
				// 跳過廠區有值
				if (!row.getValue(12208).toString().equals(""))
					continue;
				item = (IItem) row.getReferent();
				itemlist.add(item);
			}

			// Complete itemConfig
			getexcel(str_ExcelTarget, itemlist, ini, log);
			// check map
			Iterator iter = itemConfig.keySet().iterator();
			while (iter.hasNext()) {
				String key = (String) iter.next();
				// TODO arraylist datatype change
				HashMap<String, ArrayList<Cell>> val = itemConfig.get(key);
				if (val == null)
					setfailuremessage("Excel 維護錯誤 - " + key + " 找不到對應列!!!");
			}
			if (!getfailuremessage().equals(""))
				return;
			// Catch Table again
			it = itable_affectedTable.iterator();
			// loop through affected Items
			while (it.hasNext()) {
				error = false;
				// get affected item
				row = (IRow) it.next();
				// 跳過廠區有值
				if (!row.getValue(12208).toString().equals(""))
					continue;
				item = (IItem) row.getReferent();
				log.log(1, "物件: " + item);
				// parse definition for excel class
				String autoOutput = getAutoOutput(item, admin, ini, log, str_ExcelTarget);
				if (autoOutput.contains("順序1")) {
					error = true;
					continue;
				}
				// Check Error Value
				if (autoOutput.equals("")) {
					error = true;
					errorCount++;
					continue;
				}
				if (autoOutput.contains("###$$$")) {
					error = true;
					continue;
				}
				// check error
				if (!error) {
					failure = AutoSetCode.setOutputValue(row, row.getReferent(), autoOutput, ini, log,
							str_ExcelTarget, str_returnTarget);
					if (failure)
						continue;
				}
			}
		} catch (APIException e) {
			e.printStackTrace();
			throw e;
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
	}

	protected static String getAutoOutput(IItem item, IAgileSession admin, Ini ini, LogNew log, String str_ExcelTarget)
			throws APIException, Exception {
		// get agile class
		String agileAPIName = item.getAgileClass().getAPIName();
		log.log(1, "搜索" + agileAPIName + "對應的規則");
		HashMap map = itemConfig.get(agileAPIName);
//		// 取得對應分類的excel編碼資訊
		ArrayList<Cell> cellList = AutoCommonUtil.checkMinorCategory(item, map, log);
		// 根據excel編碼資訊產出編碼
		if (cellList == null) {
			setfailuremessage("Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行,順序1");
			return "Excel維護錯誤 - " + item.getAgileClass().getAPIName() + ",順序1";
		}
		String autoOutput = AutoSetCode.parseRule(cellList, item, admin, ini, log);
		log.log(1, str_ExcelTarget + ": " + autoOutput);
		return autoOutput;
	}

	private static HashMap getexcel(String str_ExcelTarget, ArrayList<IItem> itemlist, Ini ini, LogNew log)
			throws APIException {
		InputStream ExcelFileToRead = null;
		Workbook wbook = null;
		String EXCEL_FILE = "";
		log.log(">Prepare To Read Excel");
		EXCEL_FILE = ini.getValue("File Location", "EXCEL_FILE_PATH_Setting");
		try {
			ExcelFileToRead = new FileInputStream(EXCEL_FILE);
			wbook = Workbook.getWorkbook(ExcelFileToRead);
			ExcelFileToRead.close();
		} catch (FileNotFoundException e) {
			log.log("找不到該檔案，請檢查Config.ini!");
			setfailuremessage("找不到該檔案，請檢查Config.ini!");
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = null;
		sheet = wbook.getSheet(str_ExcelTarget);
		return createMap(sheet, itemlist, ini, log, str_ExcelTarget);
	}

	private static HashMap createMap(Sheet sheet, ArrayList<IItem> itemlist, Ini ini, LogNew log,
			String str_ExcelTarget) throws APIException {

		itemConfig = new HashMap<>();
		HashMap<String, ArrayList<Cell>> partData = null;
		for (int count = 0; count < itemlist.size(); count++) {

			IItem item = itemlist.get(count);
			String agileAPIName = item.getAgileClass().getAPIName();

			for (int i = 1; i < sheet.getRows(); i++) {
				Cell cell1 = sheet.getCell(1, i);// Excel API

				if (cell1.getType() == CellType.EMPTY)
					continue;
				LabelCell labelCell = (LabelCell) cell1;
				String classAPIName = labelCell.getString();

				if (classAPIName.toLowerCase().toString().equals("end"))
					break;

				Cell cell2 = sheet.getCell(2, i);// 分類依據欄位
				if (classAPIName.toLowerCase().contains((agileAPIName.toLowerCase()))) {

					if (itemConfig.get(classAPIName) != null)
						partData = itemConfig.get(classAPIName);
					else
						partData = new HashMap<String, ArrayList<Cell>>();

					// 未來會根據cell2填寫內容判斷分類
//					if (cell2.getType() != CellType.EMPTY) {
//						String cell2Data = cell2.getContents().trim();
//						ArrayList<Cell> cellList = getExcelCell(str_ExcelTarget, sheet, i, log);
//						// 空字串即無須判斷直接使用
//						if ("".equals(cell2Data))
//							partData.put("", cellList);
//						else {
//							partData.put(cell2Data, cellList);
//						}
//					} else {
					// 空字串即無須判斷直接使用
					ArrayList<Cell> cellList = getExcelCell(str_ExcelTarget, sheet, i, log);
					partData.put("", cellList);
//					}
					itemConfig.put(agileAPIName, partData);
				}
			}
		}
		return itemConfig;
	}

	private static ArrayList<Cell> getExcelCell(String str_ExcelTarget, Sheet sheet, int row, LogNew log) {

		if (row >= sheet.getRows())
			return null;

		Cell[] cellList = sheet.getRow(row);

		ArrayList<Cell> list = new ArrayList<>();
		for (int i = 2; i < cellList.length; i++) {
			if (cellList[i].getContents().trim().equals("-")) {
				continue;
			}
			if (cellList[i].equals(null) || cellList[i].getContents().trim().equals("")
					|| cellList[i].getType() == CellType.EMPTY) {
				continue;
			}
			if (cellList[i].getContents().trim().toUpperCase().equals("START")) {
				cellList[i] = sheet.getCell(i, 0);
			}
			if (cellList[i].getContents().trim().toUpperCase().equals("END")) {
				break;
			}
			list.add(cellList[i]);
		}
		for (int i = 0; i < list.size(); i++) {
			log.log(list.get(i).getContents());
		}
		return list;
	}

	public static int errorCount;
	protected IAgileSession admin;
	protected String FILEPATH;
	public static boolean failuresetname = false;// 若都沒失敗就是false
	public static String failuremessage = "";
	static boolean failure = false;
	static boolean error = false;
	static HashMap<String, HashMap<String, ArrayList<Cell>>> itemConfig;// 用於存放 Excel 中的規則

	public HashMap<String, HashMap<String, ArrayList<Cell>>> getexcelMap() {
		return itemConfig;
	}

	public boolean getfailuresetname() {
		// 若有跑錯(通常是配方) 此時會回傳true
		return failuresetname;
	}

	public static String getfailuremessage() {
		return failuremessage;
	}

	public static void setfailuremessage(String errorMessage) {
		failuremessage += errorMessage + "\n\r";
		errorCount += 1;
	}

	public static void resetfailuremessage() {
		failuremessage = "";
	}

	public void resetgetfailuresetname() {
		failuresetname = false;
	}

	public static int getErrorCount() {
		return errorCount;
	}

	public static void resetCount() {
		errorCount = 0;
	}
}
// subclass name probably can't be found due to chinese character vs apiname