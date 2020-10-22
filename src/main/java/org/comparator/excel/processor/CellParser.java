package org.comparator.excel.processor;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellParser {
	public XSSFWorkbook oldWorkbook = null;
	public XSSFWorkbook newWorkbook = null;
	LinkedList<String> uniqueKeyColumns;
	String sheetName = null;
	//(uniqueKey, RowValuesMap)
	Map<String, Map<String, String>> oldWorkbookRecords, newWorkbookRecords;

	public CellParser(Workbook oldWorkbook, Workbook newWorkbook, String sheetName, LinkedList<String> uniqueKeyColumns) {
		this.oldWorkbook = (XSSFWorkbook) oldWorkbook;
		this.newWorkbook = (XSSFWorkbook) newWorkbook;
		this.sheetName = sheetName;
		this.uniqueKeyColumns = uniqueKeyColumns;
		oldWorkbookRecords = getOldWorkBookRecords();
		newWorkbookRecords = getNewWorkBookRecords();
	}

	public Map<String, List<String>> getCommonUniqueKeys() {
		Map<String, List<String>> uniqueKeyMap = new LinkedHashMap<>();
		Row oldWorkbookHeaderRow = oldWorkbook.getSheet(sheetName).getRow(oldWorkbook.getSheet(sheetName).getFirstRowNum());
		Row newWorkbookHeaderRow = newWorkbook.getSheet(sheetName).getRow(newWorkbook.getSheet(sheetName).getFirstRowNum());
		List<String> referenceColumn = null;
		String oldWorkbookCellValue = "";
		String newWorkbookCellValue = "";
		List<String> uniqueKeyColumnList = new ArrayList<>();
		List<String> givenUniqueKeyColumnList = new ArrayList<>();
		for(String uniqueKeyColumn : uniqueKeyColumns) {
			givenUniqueKeyColumnList.add(uniqueKeyColumn.trim());
		}
		for(int oldWorkbookcell = 0; oldWorkbookcell < oldWorkbookHeaderRow.getLastCellNum(); oldWorkbookcell++) {
			oldWorkbookCellValue = oldWorkbookHeaderRow.getCell(oldWorkbookcell).getStringCellValue().trim();
			if(givenUniqueKeyColumnList.contains(oldWorkbookCellValue)) {
				for(int newWorkbookcell = 0; newWorkbookcell < newWorkbookHeaderRow.getLastCellNum(); newWorkbookcell++) {
					newWorkbookCellValue = newWorkbookHeaderRow.getCell(newWorkbookcell).getStringCellValue().trim();
					if(givenUniqueKeyColumnList.contains(newWorkbookCellValue)) {
						
						if(!uniqueKeyColumnList.contains(newWorkbookCellValue)) uniqueKeyColumnList.add(newWorkbookCellValue);
					}
				}
			}
		}
		Collections.sort(givenUniqueKeyColumnList);
		Collections.sort(uniqueKeyColumnList);
		if(givenUniqueKeyColumnList.equals(uniqueKeyColumnList)) {
			List<String> oldWorkBookUniqueKeyList = new ArrayList<>();
			int oldWorkbookRowNum = oldWorkbook.getSheet(sheetName).getLastRowNum();
			for(int row = 0; row <= oldWorkbookRowNum ; row++) {
				oldWorkbookCellValue = "";
				for(int oldWorkbookCell = 0; oldWorkbookCell < uniqueKeyColumnList.size(); oldWorkbookCell++) {
					oldWorkbookCellValue += getCellValue(oldWorkbook, sheetName, uniqueKeyColumnList.get(oldWorkbookCell), row);
				}
				if(oldWorkbookCellValue != "") oldWorkBookUniqueKeyList.add(oldWorkbookCellValue);
			}
			for(int row = 0; row <= oldWorkBookUniqueKeyList.size() ; row++) {
				referenceColumn = new LinkedList<>();
				newWorkbookCellValue = "";
				for(int newWorkbookCell = 0; newWorkbookCell < uniqueKeyColumnList.size(); newWorkbookCell++) {
					referenceColumn.add(getCellValue(newWorkbook, sheetName, uniqueKeyColumnList.get(newWorkbookCell), row));
					newWorkbookCellValue += getCellValue(newWorkbook, sheetName, uniqueKeyColumnList.get(newWorkbookCell), row);
				}
				if(!referenceColumn.isEmpty() && newWorkbookCellValue != "" && oldWorkBookUniqueKeyList.contains(newWorkbookCellValue)) uniqueKeyMap.put(newWorkbookCellValue, referenceColumn);
			}
		}
		return uniqueKeyMap;
	}
	
	public Map<String, Map<String, Map<String, String>>> getModifiedAndDeletedValues(Map<String, List<String>> uniqueKeyMap) {
		Map<String, String> oldWorkbookRowValuesMap, newWorkbookRowValuesMap;
		Map<String, Map<String, String>> unequalValuesMap = null;
		Map<String, Map<String, String>> unequalColumnValuesMap = null;
		Map<String, Map<String, Map<String, String>>> modifiedAndDeletedValuesMap = new LinkedHashMap<>();
		
		Set<Entry<String, List<String>>> uniqueKeyMapEntries = uniqueKeyMap.entrySet();
		for(Entry<String, List<String>> uniqueKeyEntry : uniqueKeyMapEntries) {
			unequalValuesMap = new LinkedHashMap<>();
			unequalColumnValuesMap = new LinkedHashMap<>();
			oldWorkbookRowValuesMap = oldWorkbookRecords.get(uniqueKeyEntry.getKey());
			newWorkbookRowValuesMap = newWorkbookRecords.get(uniqueKeyEntry.getKey());
			if(oldWorkbookRecords.keySet().contains(uniqueKeyEntry.getKey()) && newWorkbookRecords.keySet().contains(uniqueKeyEntry.getKey())) {
				if(!areEqualKeyValues(oldWorkbookRowValuesMap, newWorkbookRowValuesMap)) {
					unequalValuesMap = getUnequalAndDeletedValues(oldWorkbookRowValuesMap, newWorkbookRowValuesMap);
				}
				if(!oldWorkbookRowValuesMap.equals(newWorkbookRowValuesMap)) {
					unequalColumnValuesMap = getUnequalColumnValues(oldWorkbookRowValuesMap, newWorkbookRowValuesMap);
					modifiedAndDeletedValuesMap.put(uniqueKeyEntry.getKey(), unequalColumnValuesMap);
				}
			}
			if(!unequalValuesMap.isEmpty()) modifiedAndDeletedValuesMap.put(uniqueKeyEntry.getKey(), unequalValuesMap);
		}
		return modifiedAndDeletedValuesMap;
	}

	private Map<String, Map<String, String>> getUnequalColumnValues(Map<String, String> oldWorkbookRowValuesMap,
			Map<String, String> newWorkbookRowValuesMap) {
		Map<String, Map<String, String>> unequalColumnValuesMap = new LinkedHashMap<>();
		Map<String, String> oldNewValuesMap = null;
		Set<Entry<String, String>> newWorkbookValuesMapEntries = newWorkbookRowValuesMap.entrySet();
		for(Entry<String, String> newWorkbookValuesMapEntry : newWorkbookValuesMapEntries) {
			String key = "";
			if(!oldWorkbookRowValuesMap.keySet().contains(newWorkbookValuesMapEntry.getKey())) {
				if(!newWorkbookValuesMapEntry.getValue().equals(oldWorkbookRowValuesMap.get(newWorkbookValuesMapEntry.getKey()))) {
					key = newWorkbookValuesMapEntry.getKey();
					oldNewValuesMap = new LinkedHashMap<>();
					if(oldWorkbookRowValuesMap.get(newWorkbookValuesMapEntry.getKey()) == null) oldNewValuesMap.put("", newWorkbookValuesMapEntry.getValue());
					else oldNewValuesMap.put(oldWorkbookRowValuesMap.get(newWorkbookValuesMapEntry.getKey()), newWorkbookValuesMapEntry.getValue());
					oldNewValuesMap.remove("","");
					oldNewValuesMap.remove("",null);
					oldNewValuesMap.remove(null,null);
					oldNewValuesMap.remove(null,null);
				}
			}
			if(key !="" && oldNewValuesMap != null && !oldNewValuesMap.isEmpty()) unequalColumnValuesMap.put(key, oldNewValuesMap);
		}
		return unequalColumnValuesMap;
	}

	public Map<String, Map<String, String>> getInsertedRows(Map<String, List<String>> uniqueKeyMap) {
		Map<String, Map<String, String>> uniqueOldWorkbookRecords = oldWorkbookRecords.entrySet().stream()
				.filter(entry -> !uniqueKeyMap.keySet().contains(entry.getKey()))
				.collect(Collectors.toMap(entry -> entry.getKey(), entry -> entry.getValue()));
		Map<String, Map<String, String>> uniqueNewWorkbookRecords = newWorkbookRecords.entrySet().stream()
				.filter(entry -> !uniqueKeyMap.keySet().contains(entry.getKey()))
				.collect(Collectors.toMap(entry -> entry.getKey(), entry -> entry.getValue()));
		
		Map<String, Map<String, String>> insertedRowsMap = uniqueNewWorkbookRecords.entrySet().stream()
				.filter(entry -> !uniqueOldWorkbookRecords.keySet().contains(entry.getKey()))
				.filter(entry -> !entry.getKey().isEmpty())
				.collect(Collectors.toMap(entry -> entry.getKey(), entry -> entry.getValue()));
		
		return insertedRowsMap;

	}
	
	public Map<String, Map<String, String>> getDeletedRows(Map<String, List<String>> uniqueKeyMap) {
		Map<String, Map<String, String>> uniqueOldWorkbookRecords = oldWorkbookRecords.entrySet().stream()
				.filter(entry -> !uniqueKeyMap.keySet().contains(entry.getKey()))
				.collect(Collectors.toMap(entry -> entry.getKey(), entry -> entry.getValue()));
		Map<String, Map<String, String>> uniqueNewWorkbookRecords = newWorkbookRecords.entrySet().stream()
				.filter(entry -> !uniqueKeyMap.keySet().contains(entry.getKey()))
				.collect(Collectors.toMap(entry -> entry.getKey(), entry -> entry.getValue()));
		
		Map<String, Map<String, String>> deletedRowsMap = uniqueOldWorkbookRecords.entrySet().stream()
				.filter(entry -> !uniqueNewWorkbookRecords.keySet().contains(entry.getKey()))
				.filter(entry -> !entry.getKey().isEmpty())
				.collect(Collectors.toMap(entry -> entry.getKey(), entry -> entry.getValue()));
		
		return deletedRowsMap;

	}

	private Map<String, Map<String, String>> getUnequalAndDeletedValues(Map<String, String> oldWorkbookRowValuesMap,
			Map<String, String> newWorkbookRowValuesMap) {
		Map<String, Map<String, String>> unequalKeyValuesMap = new LinkedHashMap<>();
		Map<String, String> oldNewValuesMap = null;
		Set<Entry<String, String>> oldWorkbookValuesMapEntries = oldWorkbookRowValuesMap.entrySet();
		for(Entry<String, String> oldWorkbookValuesMapEntry : oldWorkbookValuesMapEntries) {
			String key = "";
			if(newWorkbookRowValuesMap.keySet().contains(oldWorkbookValuesMapEntry.getKey())) {
				if(!oldWorkbookValuesMapEntry.getValue().equals(newWorkbookRowValuesMap.get(oldWorkbookValuesMapEntry.getKey()))) {
					key = oldWorkbookValuesMapEntry.getKey();
					oldNewValuesMap = new LinkedHashMap<>();
					oldNewValuesMap.put(oldWorkbookValuesMapEntry.getValue(), newWorkbookRowValuesMap.get(oldWorkbookValuesMapEntry.getKey()));
				}
			}
			if(key !="" && oldNewValuesMap != null && !oldNewValuesMap.isEmpty()) {
				unequalKeyValuesMap.put(key, oldNewValuesMap);
			} 
		} 
		return unequalKeyValuesMap;
	}

	private boolean areEqualKeyValues(Map<String, String> oldWorkbookValuesMap,
			Map<String, String> newWorkbookValuesMap) {
		return oldWorkbookValuesMap.entrySet().stream().allMatch(e -> 
					e.getValue().equals(newWorkbookValuesMap.get(e.getKey())));
	}

	private Map<String, Map<String, String>> getNewWorkBookRecords() {
		Map<String, Map<String, String>> newWorkbookRecords = new LinkedHashMap<String, Map<String,String>>();
		Map<String, String> newWorkbookrowValuesMap = null;
		Set<String> uniqueKeyValuesSet = null;
		int newWorkbookRowNum = newWorkbook.getSheet(sheetName).getLastRowNum();
		for(int row = 0; row <= newWorkbookRowNum; row++) {
			String uniqueKey = "";
			uniqueKeyValuesSet = new LinkedHashSet<>();
			newWorkbookrowValuesMap = new LinkedHashMap<>();
			String newWorkbookCellValue = "";
			int columnSize = newWorkbook.getSheet(sheetName).getRow(newWorkbook.getSheet(sheetName).getFirstRowNum()).getLastCellNum();
			for(int cell = 0; cell < columnSize; cell++) {
				String newWorkbookColumnName = newWorkbook.getSheet(sheetName).getRow(newWorkbook.getSheet(sheetName).getFirstRowNum()).getCell(cell).getStringCellValue();
				newWorkbookCellValue = getCellValue(newWorkbook, sheetName, newWorkbookColumnName, row);
				newWorkbookrowValuesMap.put(newWorkbookColumnName, newWorkbookCellValue);
				if(uniqueKeyColumns.contains(newWorkbookColumnName)) {
					uniqueKeyValuesSet.add(newWorkbookCellValue);
				}
			}
			for(String key : uniqueKeyValuesSet) {
				uniqueKey += key;
			}
			newWorkbookRecords.put(uniqueKey, newWorkbookrowValuesMap);
		}
		return newWorkbookRecords;
	}

	private Map<String, Map<String, String>> getOldWorkBookRecords() {
		Map<String, Map<String, String>> oldWorkbookRecords = new LinkedHashMap<String, Map<String,String>>();
		Map<String, String> oldWorkbookRowValuesMap = null;
		Set<String> uniqueKeyValuesSet = null;
		int oldWorkbookRowNum = oldWorkbook.getSheet(sheetName).getLastRowNum();
		for(int row = 0; row <= oldWorkbookRowNum; row++) {
			String uniqueKey = "";
			uniqueKeyValuesSet = new LinkedHashSet<>();
			oldWorkbookRowValuesMap = new LinkedHashMap<>();
			String oldWorkbookCellValue = "";
			int columnSize = oldWorkbook.getSheet(sheetName).getRow(oldWorkbook.getSheet(sheetName).getFirstRowNum()).getLastCellNum();
			for(int cell = 0; cell < columnSize; cell++) {
				String oldWorkbookColumnName = oldWorkbook.getSheet(sheetName).getRow(oldWorkbook.getSheet(sheetName).getFirstRowNum()).getCell(cell).getStringCellValue();
				oldWorkbookCellValue = getCellValue(oldWorkbook, sheetName, oldWorkbookColumnName, row);
				oldWorkbookRowValuesMap.put(oldWorkbookColumnName, oldWorkbookCellValue);
				if(uniqueKeyColumns.contains(oldWorkbookColumnName)) {
					uniqueKeyValuesSet.add(oldWorkbookCellValue);
				}
			}
			for(String key : uniqueKeyValuesSet) {
				uniqueKey += key;
			}
			oldWorkbookRecords.put(uniqueKey, oldWorkbookRowValuesMap);
		}
		return oldWorkbookRecords;
	}
	
	private String getCellValue(XSSFWorkbook workbook, String sheetName, String columnName, int rowNum) {
		try {
            int columnNum = -1;
            XSSFSheet sheet = workbook.getSheet(sheetName);
            XSSFRow row = sheet.getRow(0);
            XSSFCell cell = null;
            for(int i = 0; i < row.getLastCellNum(); i++) {
                if(row.getCell(i).getStringCellValue().trim().equals(columnName.trim()))
                	columnNum = i;
            }
 
            row = sheet.getRow(rowNum);
            if(row != null) cell = row.getCell(columnNum);
            else return "";
 
            if(cell != null) {
            	if(cell.getCellType() == CellType.STRING) return cell.getStringCellValue();
                else if(cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA) {
                	String cellValue = null;
                	if(cell.getNumericCellValue() == (int)cell.getNumericCellValue()) cellValue = String.valueOf((int)cell.getNumericCellValue());
                	else if(DateUtil.isCellDateFormatted(cell)) {
                    	DataFormatter df = new DataFormatter();
                        cellValue = df.formatCellValue(cell);
                    } else cellValue = String.valueOf(cell.getNumericCellValue());
                    return cellValue;
                }else if(cell.getCellType() == CellType.BLANK)
                    return "";
                else
                    return String.valueOf(cell.getBooleanCellValue());
            } else return "";
        }
        catch(Exception e) {
            e.printStackTrace();
            return "";
        }
    }
}
