import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.IChange;
import com.agile.api.IDataObject;
import com.agile.api.INode;
import com.agile.api.IRoutable;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.ICustomAction;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;
import com.agile.px.IWFChangeStatusEventInfo;

import jxl.Cell;
import jxl.LabelCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import AutoOutput.CombinedOutput;
import AutoOutput.AutoSetCode;
import record.LogNew;
import util.AUtil;
import util.Ini;

public class Modularization_AutoNumberOrDescriptionOrStandard implements ICustomAction,IEventAction{

	/**********************************************
	 * 主程式執行 - PX
	 **********************************************/
	@Override
	public ActionResult doAction(IAgileSession iAgileSession, INode iNode, IDataObject change) {
		Ini ini = new Ini("D:\\Anselm_Program_Data\\Config.ini");
		LogNew log = new LogNew("AutoNumberOrDescriptionOrStandard");
		String FILEPATH = ini.getValue("Program Use", "LogFile") + "AutoNumberOrDescriptionOrStandard_";
		IAgileSession admin = AUtil.getAgileSession(ini, "AgileAP");	
		log.logSeparatorBar();
		String str_resultCode = "";
		String[] str_list;
		String[] str_list1;
		String str_result = "";
		String str_status = null;
		String str_workflow =null;
		int errorCountCode = 0;
		IChange ichange_changeOrder = null;
		
		try {
			log.setLogFile(FILEPATH+change+".log");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		try {
			ichange_changeOrder = (IChange) admin.getObject(IChange.OBJECT_TYPE, change.getName());
			str_status = ichange_changeOrder.getStatus().getAPIName();
			str_workflow = ichange_changeOrder.getWorkflow().getAPIName();
		} catch (Exception e) {
			log.log(e);
			str_result = "Object is not Change!";
			close(ini, log);
			return new ActionResult(ActionResult.STRING, str_result);
		}
		
		try {
			log.log(admin.getCurrentUser());
		} catch (APIException e) {
			log.log(e);
			str_result = "getCurrentUser()發生錯誤";
			close(ini, log);
			return new ActionResult(ActionResult.STRING, str_result);
		}
		
		try {
			str_result = getExcelSetting(log, ini,ichange_changeOrder,admin);
		} catch (APIException e) {
			log.log(e);
			str_result = "getExcelSetting()呼叫發生錯誤";
			close(ini, log);
			return new ActionResult(ActionResult.STRING, str_result);
		}
		
		if(str_result.equals("找不到該檔案，請檢查Config.ini!")) {
			str_result = "找不到該檔案，請檢查Config.ini!";
			close(ini, log);
			return new ActionResult(ActionResult.STRING,str_result );
		}
		if(str_result.contains("找不到該流程，請檢查Excel!")) {
			str_result = "找不到該流程，請檢查Excel!";
			close(ini, log);
			return new ActionResult(ActionResult.STRING,str_result);
		}		
		str_list= str_result.split("##");
		for (int i = 0; i < str_list.length; i++) {
			str_list1 = str_list[i].split(":");
			if(str_list1[0].equals("end")) {
				break;
			}
			if(!str_list1[0].equals(str_status)){
				continue;
			}
			if(str_list1[0].length()<3) {
				return new ActionResult(ActionResult.STRING,"Excel維護錯誤 - " + str_workflow + "行" + ",參數" + i + "");
			}
			String str_excelTarget  =str_list1[2];
			String str_returnTarget = str_list1[1];
			if(str_list1[0].equals(str_status)) {
				try {
					// Auto Code
					str_result="";
					AutoSetCode code = new AutoSetCode();
					code.action(ichange_changeOrder, admin, ini, log,str_excelTarget,str_returnTarget);
					errorCountCode = CombinedOutput.getErrorCount();
					str_result = errorCountCode == 0 ? "程式執行完成" : "執行自動"+str_excelTarget+"總共有" + errorCountCode + "筆條件失敗，請檢查log檔" + "\n\r";
				} catch (Exception e) {
					close(ini, log);
					return new ActionResult(ActionResult.STRING, str_result);
				}
				if (errorCountCode != 0) {
					str_result += CombinedOutput.getfailuremessage();
					CombinedOutput.resetCount();
					CombinedOutput.resetfailuremessage();
					close(ini, log);
					return new ActionResult(ActionResult.STRING, str_result);
				}
			}
		}
		close(ini, log);
		return new ActionResult(ActionResult.STRING,"done");
	}
	/**********************************************
	 * 主程式執行 - Event
	 **********************************************/
	@Override
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo req) {
		Ini ini = null;
		LogNew log = null;
		IAgileSession admin = null;
		String FILEPATH = null;
		log.logSeparatorBar();
		IWFChangeStatusEventInfo info = (IWFChangeStatusEventInfo) req;
		String result = "";
		try {
			log.setLogFile(FILEPATH);
		} catch (IOException e) {
			e.printStackTrace();
		}
		log.logSeparatorBar();
		IChange changeOrder = null;
		try {
			changeOrder = (IChange) info.getDataObject();
		} catch (APIException e) {
			log.log(e);
			result = "Object is not Change !!!";
			close(ini, log);
			return new EventActionResult(req, new ActionResult(ActionResult.STRING, result));
		}
		return new EventActionResult(req, new ActionResult(ActionResult.STRING, result));
	}

	/**********************************************
	 * 關閉建構
	 * @param log 
	 * @param ini 
	 **********************************************/
	public void close(Ini ini, LogNew log) {
		ini = null;
		try {
			if (log != null)
				log.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**********************************************
	 * 得到Excel設定資訊
	 * @param ichange_changeOrder 
	 * @param admin 
	 * @return 
	 * @throws APIException 
	 * 
	 * 
	 **********************************************/
	private String getExcelSetting(LogNew log,Ini ini, IChange ichange_changeOrder, IAgileSession admin) throws APIException {
		// TODO Auto-generated method stub
		InputStream ExcelFileToRead = null;
		Workbook wbook = null;
		String EXCEL_FILE = "";
		String str_result = "";
		String[] str_list;
		String str_workflow =null;
		String str_status = null;
		str_workflow = ichange_changeOrder.getWorkflow().getAPIName();
		str_status = ichange_changeOrder.getStatus().getAPIName();
		log.log(">Prepare To Read Excel Setting");
		EXCEL_FILE = ini.getValue("File Location", "EXCEL_FILE_PATH_Setting");
		try {
			ExcelFileToRead = new FileInputStream(EXCEL_FILE);
			wbook = Workbook.getWorkbook(ExcelFileToRead);
			ExcelFileToRead.close();
		} catch (FileNotFoundException e) {
			log.log("找不到該檔案，請檢查Config.ini!");
			return "找不到該檔案，請檢查Config.ini!";
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = null;
		sheet = wbook.getSheet("設定");
		for (int i = 1; i < sheet.getRows(); i++) {
			Cell cell = sheet.getCell(1, i);// Workflow
			LabelCell labelCell = (LabelCell) cell;
			if(str_workflow.equals(labelCell.getString())) {
				log.log("當前工作流程："+str_workflow);
				log.log("當前站別："+str_status);
				for (int j = 2; j < sheet.getRows(); j++) {
					Cell cell1 = sheet.getCell(j, i);
					str_result += cell1.getContents()+"##";
					if(cell1.getContents().equals("end")) {
						break;
					}
				}
				break;
			}
			if(labelCell.getString().toLowerCase().toString().equals("end")) {
				log.log("找不到該流程，請檢查Excel!");
				return "找不到該流程，請檢查Excel!";
			}
		}		
		return str_result;
	}
}

