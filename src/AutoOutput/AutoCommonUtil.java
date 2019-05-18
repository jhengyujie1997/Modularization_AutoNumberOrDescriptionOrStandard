package AutoOutput;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

import com.agile.api.CommonConstants;
import com.agile.api.IAgileList;
import com.agile.api.ICell;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;

import jxl.Cell;
import record.LogNew;

public class AutoCommonUtil {
	public static  ArrayList<Cell> checkMinorCategory(IItem item,Map map,LogNew log) throws Exception {	
		ArrayList<Cell> cellList = new ArrayList<Cell>();
		ArrayList<Cell> tmp = new ArrayList<Cell>();
		for (Object key : map.keySet()) {
            String keyS = key.toString();
            // 表示無須根據任何分類判斷，直接回傳即可
            if("".equals(keyS) || "-".equals(keyS)) {
            	cellList = (ArrayList<Cell>) map.get(key);
            	break;
            }
            String[] splitKey = keyS.split(",");
            if(splitKey.length<3) {
            	log.log("Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行,順序1");
            	return cellList;
            }
            ITable columnTable = null;
            if(splitKey[1].equalsIgnoreCase("P3"))
            	columnTable = item.getTable(CommonConstants.TABLE_PAGETHREE);
            else if(splitKey[1].equalsIgnoreCase("P2"))
            	columnTable = item.getTable(CommonConstants.TABLE_PAGETWO);
            else if(splitKey[1].equalsIgnoreCase("tb"))
            	columnTable = item.getTable(ItemConstants.TABLE_TITLEBLOCK);
            Iterator it = columnTable.iterator();
            IRow row = (IRow) it.next();
            if("".equalsIgnoreCase(splitKey[2].trim())) {
            	log.log("Excel維護錯誤 - " + item.getAgileClass().getAPIName() + "行,順序1");
            	return cellList;
            }
            // 除判斷值以外內容都代入Others
            if("Others".equalsIgnoreCase(splitKey[2].trim())) {
            	tmp = (ArrayList<Cell>) map.get(key);
            	continue;
            }
            ICell cell = row.getCell(splitKey[0]);
            IAgileList list = (IAgileList) cell.getValue();
            IAgileList[] selected = list.getSelection();
            if(selected!=null) {
            	String targetValue = selected[0].getAPIName();
	            if(targetValue.trim().equals(splitKey[2].trim())) {
	            	log.log(2,"取得小分類: "+selected[0].getValue());
	            	log.log(2,"小分類 API Name: "+targetValue);
	            	cellList = (ArrayList<Cell>) map.get(key);
	            }
            }
            
        }
		if(cellList==null) {
			log.log(2,"無對應小分類，代入預設分類");
			cellList = tmp;
		}
		return cellList;
	}
}
