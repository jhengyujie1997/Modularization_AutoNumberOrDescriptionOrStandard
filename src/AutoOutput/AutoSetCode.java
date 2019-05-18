package AutoOutput;

import java.util.ArrayList;
import java.util.Iterator;

import com.agile.api.APIException;
import com.agile.api.ChangeConstants;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.IQuery;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.NumberCell;
import record.LogNew;
import util.Ini;

public class AutoSetCode {
	public void action(IChange ichange_changeOrder, IAgileSession admin, Ini ini, LogNew log, String str_excelTarget,
			String str_returnTarget) throws Exception {
		log.log(1, "Set " + str_excelTarget + "：......");
		IChange changeAF = (IChange) admin.getObject(IChange.OBJECT_TYPE, ichange_changeOrder.getName());
		ITable AffectedItemsTable = changeAF.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
		Iterator it = AffectedItemsTable.iterator();
//		if (str_excelTarget.equalsIgnoreCase("編碼")) {
//			while (it.hasNext()) {
//				IRow row = (IRow) it.next();
//				IDataObject item = row.getReferent();
//				String list[] = str_returnTarget.split(",");
//				log.log("ExcelTarget:" + list[0]);
//				log.log("頁籤:" + list[1]);
//				if (list[0].substring(0, 1).equals("@")) {
//					ITable redLineTitleBlock = null;
//					if (list[1].toString().toUpperCase().equals("TITLEBLOCK")) {
//						redLineTitleBlock = item.getTable(ItemConstants.TABLE_REDLINETITLEBLOCK);
//					} else if (list[1].toString().toUpperCase().equals("P2")) {
//						redLineTitleBlock = item.getTable(ItemConstants.TABLE_REDLINEPAGETWO);
//					} else if (list[1].toString().toUpperCase().equals("P3")) {
//						redLineTitleBlock = item.getTable(ItemConstants.TABLE_REDLINEPAGETHREE);
//					}
//					IRow redlineRow = (IRow) redLineTitleBlock.iterator().next();
//					ICell cell = redlineRow.getCell(list[0].substring(1));
//					if (list[0].substring(1).equals("number")) {
//						cell.setValue(cell.getOldValue());
//					} else if (redlineRow.getCell(list[0].substring(1)).toString() == "") {
//						cell.setValue("");
//					} else {
//						redlineRow.getCell(list[0].substring(1)).setValue("");
//					}
//				} else {
//					ICell cell = item.getCell(list[0]);
//					cell.setValue(cell.getOldValue());
//				}
//			}
//		}
//		log.log("BBB");
		CombinedOutput.start(ichange_changeOrder, admin, ini, log, str_excelTarget, str_returnTarget);
	}

	public static String parseRule(ArrayList<Cell> cellList, IItem item, IAgileSession admin, Ini ini, LogNew log)
			throws Exception {

		if (cellList == null) {
			log.log(1, "找不到規則" + "!!!");
			CombinedOutput.setfailuremessage(item.getName() + "找不到規則 !!!");
			return "";
		}
		String autoCode = "";
		// Get target row's cells

		try {
			String error = "###$$$";// 回傳是否有錯誤
			for (int i = 1; i < cellList.size(); i++) {
				Cell cell = cellList.get(i);
				int cellint = Integer.parseInt(cellList.get(0).getContents().toString().substring(
						cellList.get(0).getContents().toString().length() - 1,
						cellList.get(0).getContents().toString().length()));
				String celli = "順序" + (cellint + i);// 順序從cellList.get(0)開始

				if (cell.getType() == CellType.EMPTY)
					continue;
				log.log(1, celli + " 規則為" + cell.getContents());
				if (cell.getType() == CellType.LABEL) {
					LabelCell labelCell = (LabelCell) cell;
					String c = labelCell.getString().trim();

					if (c.toLowerCase().equals("end"))
						break;
					else if (c.charAt(0) == '$') {
						if (c.length() < 2) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "," + celli + "");
						}
						if (!c.contains("nbsp"))
							autoCode += c.substring(1).replace(" ", "") + "";
						else
							autoCode += " ";
					} // end else if - '$'

					else if (c.charAt(0) == '@') {
						if (!c.contains(",")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "," + celli + "");
						}
						String[] rules = c.substring(1).split(",");
						if (rules.length != 3) {
							if (rules == null) {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
								continue;
							} else {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
								continue;
							}
						}
						char c0 = rules[1].toLowerCase().charAt(0);
						char c1 = rules[1].toLowerCase().charAt(1);
						if (c0 != 'p') {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						if (AutoGetInf.check(c1)) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						c0 = rules[2].toLowerCase().charAt(0);
						if (AutoGetInf.check(c0)) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}

						String ObjectInfo = AutoGetInf.getListValue(item, log, admin, rules[0], rules[1], rules[2]);

						if (ObjectInfo.contains("欄位為空")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else if (ObjectInfo.contains("請檢查 Excel")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else if (ObjectInfo.contains("維護錯誤")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else
							autoCode += ObjectInfo;

					} // end else if - '@'
					else if (c.charAt(0) == '~') {
						// 流水碼
						if (!c.contains(",")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						String[] rules = c.substring(1).split(",");
						if (rules[0].contains("流水碼")) {
							int rule = Integer.parseInt(rules[1]);

							// Initialize serialNumber
							String serialNumber = initializeNum(rule);
							while (checkSameNum(autoCode + serialNumber, item, admin, ini, log)) {
								int serialInt = Integer.parseInt(serialNumber) + 1;
								serialNumber = String.valueOf(serialInt);
								serialNumber = addZero(rule, serialNumber);
								if (checkOverNum(serialNumber, rule)) {
									autoCode += error;
									CombinedOutput.setfailuremessage("流水號已滿... 請洽 Admin!!!");
								}
							}
							autoCode += addZero(rule, serialNumber);
						}
					} // end else if - '~'
					else if (c.charAt(0) == '&') {
						String Code = c.substring(1, c.length()).replace(" ", "");
						log.log(1, "讀取 Config ...");
						String string = ini.getValue("Excel-BookMark", Code);
						String[] config = string.split("/");
						String ObjectInfo = AutoGetInf.getObjectValue(item, log, admin, config);
						if (ObjectInfo.contains("欄位為空"))
							CombinedOutput.setfailuremessage(ObjectInfo);
						else
							autoCode += ObjectInfo;
					} // end else if - '&'
					else if (c.charAt(0) == '*') {
						if (!c.contains(",")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "," + celli + "");
						}
						String[] rules = c.substring(1).split(",");
						if (rules.length != 2) {
							if (rules == null) {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
								continue;
							} else {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
								continue;
							}
						}
						char c0 = rules[1].toLowerCase().charAt(0);
						char c1 = rules[1].toLowerCase().charAt(1);
						if (c0 != 'p') {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						if (AutoGetInf.check(c1)) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						String ObjectInfo = AutoGetInf.getcolunmValue(item, log, admin, rules[0], rules[1], null);

						if (ObjectInfo.contains("欄位為空")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else if (ObjectInfo.contains("請檢查 Excel")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else if (ObjectInfo.contains("維護錯誤")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else { // 去除小數點後多餘的0 => 4.0 變成 4 , 4.20 變成 4.2
							double doubles = Double.parseDouble(ObjectInfo);
							int ints = (int) doubles;
							if (ints == doubles) {
								ObjectInfo = String.valueOf(ints);
							} else {
								ObjectInfo = String.valueOf(doubles);
							}
							autoCode += ObjectInfo;
						}

					} // end else if - '*'
					else {
						String ObjectInfo = "";
						if (!c.contains(",")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(
									"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
						}
						String[] rules = c.split(",");

						if (rules.length == 3) {
							char c0 = rules[1].toLowerCase().charAt(0);
							char c1 = rules[1].toLowerCase().charAt(1);
							char c2 = rules[2].toLowerCase().charAt(0);
							if (c0 != 'p') {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							if (AutoGetInf.check(c1)) {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							if (AutoGetInf.check(c2)) {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							ObjectInfo = AutoGetInf.getcolunmValue(item, log, admin, rules[0], rules[1], rules[2]);
						} else if (rules.length == 2) {
							char c0 = rules[1].toLowerCase().charAt(0);
							char c1 = rules[1].toLowerCase().charAt(1);
							if (c0 != 'p') {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							if (AutoGetInf.check(c1)) {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
							ObjectInfo = AutoGetInf.getcolunmValue(item, log, admin, rules[0], rules[1], null);
						} else {
							if (rules == null) {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							} else {
								autoCode += error;
								CombinedOutput.setfailuremessage(
										"Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行" + "," + celli + "");
							}
						}
						if (ObjectInfo.contains("欄位為空")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else if (ObjectInfo.contains("請檢查 Excel")) {
							autoCode += error;
							CombinedOutput.setfailuremessage(ObjectInfo);
						} else {
							autoCode += ObjectInfo.trim();
							if (ObjectInfo.equals("")
									&& autoCode.substring(autoCode.length() - 1).matches("\\W") == true) {
								autoCode = autoCode.substring(0, autoCode.length() - 1);
							}
						}
					} // end else - no tag
				} // end if - Label
			} // end for
		} catch (APIException e) {
			e.printStackTrace();
		}
		return autoCode;
	}

	private static boolean checkSameNum(String number, IItem item, IAgileSession admin, Ini ini, LogNew logger)
			throws APIException {
		try {
			logger.log(">Checking Same Number");
			IQuery query = (IQuery) admin.createObject(IQuery.OBJECT_TYPE, item.getAgileClass().getAPIName());
			query.setCaseSensitive(false);
			// 1001 = item number , 2011 = P2 text05
			query.setCriteria("[1001] starts with '" + number + "' or [2011] starts with '" + number + "'");
			ITable results = query.execute();
			if (results.size() > 0)
				return true;
			else
				return false;
		} catch (APIException apie) {
			apie.printStackTrace();
			logger.logException(apie);
			throw apie;
		} catch (Exception e) {
			e.printStackTrace();
			logger.logException(e);
			throw e;
		}
	}

	private static String initializeNum(int rule) {
		int initialize = rule - 1;
		String initializeS = "";
		for (int i = 0; i < initialize; i++)
			initializeS += "0";
		return initializeS + "1";
	}

	private static String addZero(int length, String number) {
		String tmp = "";
		for (int i = 0; i < length - number.length(); i++) {
			tmp += "0";
		}
		return tmp + number;
	}

	private static boolean checkOverNum(String serialNumber, int rule) {
		int initialize = rule;
		String initializeS = "";
		for (int i = 0; i < initialize; i++)
			initializeS += "9";
		if (serialNumber.equals(initializeS))
			return true;
		else
			return false;
	}

	public static boolean setOutputValue(IRow row, IDataObject ido, String autoOutput, Ini ini, LogNew log,
			String str_ExcelTarget, String str_returnTarget) throws APIException {
		log.log(">Reset：" + ido.getName() + " " + str_returnTarget + "...");
		if (str_ExcelTarget.equals("品名")) {
			try {
				if (str_returnTarget.substring(0, 1).equals("@")) {
					ITable redLinePageTwo = ido.getTable(ItemConstants.TABLE_REDLINEPAGETWO);
					IRow redlineRow = (IRow) redLinePageTwo.iterator().next();
					if (redlineRow.getCell(str_returnTarget.substring(1)) == null) {
						CombinedOutput.failuresetname = true;
						CombinedOutput.failuremessage += ido.getName() + "無 " + str_ExcelTarget + " 欄位!!";
					} else {
						if (autoOutput.contains("###$$$"))
							autoOutput = autoOutput.replaceAll("###$$$", "");
						ICell cell = redlineRow.getCell(str_returnTarget.substring(1));
						cell.setValue(autoOutput);
						CombinedOutput.failuresetname = false;
					}
				} else {
					if (row.getCell(str_returnTarget) == null) {
						CombinedOutput.failuresetname = true;
						CombinedOutput.failuremessage += ido.getName() + "無 " + str_ExcelTarget + " 欄位!!";
					} else {
						if (autoOutput.contains("###$$$"))
							autoOutput = autoOutput.replaceAll("###$$$", "");
						ICell cell = row.getCell(str_returnTarget);
						cell.setValue(autoOutput);
						CombinedOutput.failuresetname = false;
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
				CombinedOutput.failuresetname = true;
				CombinedOutput.errorCount++;
				CombinedOutput.failuremessage += CombinedOutput.failuremessage + " " + e.getMessage();
			}
		} else if (str_ExcelTarget.equals("規格")) {
			try {
				if (str_returnTarget.substring(0, 1).equals("@")) {
					ITable redLinePageTwo = ido.getTable(ItemConstants.TABLE_REDLINEPAGETWO);
					IRow redlineRow = (IRow) redLinePageTwo.iterator().next();
					if (redlineRow.getCell(str_returnTarget.substring(1)) == null) {
						CombinedOutput.failuresetname = true;
						CombinedOutput.failuremessage += ido.getName() + "無規格欄位!!";
					} else {
						if (autoOutput.contains("###$$$"))
							autoOutput = autoOutput.replaceAll("###$$$", "");
						ICell cell = redlineRow.getCell(str_returnTarget.substring(1));
						if (!autoOutput.contains("###$$$")) {
							cell.setValue(autoOutput);
							CombinedOutput.failuresetname = false;
						}
					}
				} else {
					IItem item = (IItem) ido;
					item.setValue(str_returnTarget, autoOutput.toUpperCase());
					CombinedOutput.failuresetname = false;
				}
			} catch (Exception e) {
				CombinedOutput.failuresetname = true;
				CombinedOutput.errorCount++;
				CombinedOutput.failuremessage += CombinedOutput.failuremessage + " " + e.getMessage();
			}
		} else if (str_ExcelTarget.equals("編碼")) {
			try {
				if (str_returnTarget.substring(0, 1).equals("@")) {
					ITable redLineTitleBlock = ido.getTable(ItemConstants.TABLE_REDLINETITLEBLOCK);
					IRow redlineRow = (IRow) redLineTitleBlock.iterator().next();
					ICell cell = redlineRow.getCell(str_returnTarget.substring(1));
					cell.setValue(autoOutput.toUpperCase());
					CombinedOutput.failuresetname = false;
				} else {
					IItem item = (IItem) ido;
					item.setValue(str_returnTarget, autoOutput.toUpperCase());
					CombinedOutput.failuresetname = false;
				}
			} catch (Exception e) {
				CombinedOutput.failuresetname = true;
				CombinedOutput.errorCount++;
				CombinedOutput.failuremessage = CombinedOutput.failuremessage + " " + e.getMessage();
			}
		}
		return false;
	}
}