package org.comparator.excel.processor;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RowParser {

	/**
	 * @param oldWorkbook
	 * @param newWorkbook
	 * @param sheetName
	 * @param uniqueKeyColumns
	 */
	public void compareRows(Workbook oldWorkbook, Workbook newWorkbook, String sheetName,
			LinkedList<String> uniqueKeyColumns) {
		CellParser cellParser = new CellParser(oldWorkbook, newWorkbook, sheetName, uniqueKeyColumns);
		
		Map<String, Map<String, String>> uniqueKeyMap = cellParser.getCommonUniqueKeys();
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet( " Data ");
		XSSFRow row;
			
			Map<String, Map<String, String>> addedRecordsMap = cellParser.getAddedRecords(uniqueKeyMap);
			int rowid = 0, cellid = 0;
			if(!addedRecordsMap.isEmpty() && addedRecordsMap != null) {
				spreadsheet.createRow(rowid++).createCell(cellid).setCellValue("Added Records");
				Set<String> firstKeySet = addedRecordsMap.values().stream().findFirst().get().keySet();
				cellid = 0;
				row = spreadsheet.createRow(rowid++);
				for(String columnName : firstKeySet) {
					Cell cell = row.createCell(cellid++);
			        cell.setCellValue(columnName);
				}
				for(Entry<String, Map<String, String>> addedRecordsMapEntry : addedRecordsMap.entrySet()) {
					row = spreadsheet.createRow(rowid++);
					cellid = 0;
					for(Entry<String, String> addedRecordsMapCellEntry : addedRecordsMapEntry.getValue().entrySet()) {
						if(firstKeySet.equals(addedRecordsMapEntry.getValue().keySet())) {
							Cell cell = row.createCell(cellid++);
							cell.setCellValue(addedRecordsMapCellEntry.getValue());
						}
					}
				}
			}
			spreadsheet.createRow(rowid++);
			Map<String, Map<String, String>> deletedRecordsMap = cellParser.getDeletedRecords(uniqueKeyMap);
			if(!deletedRecordsMap.isEmpty() && deletedRecordsMap != null) {
				Set<String> firstKeySet = deletedRecordsMap.values().stream().findFirst().get().keySet();
				cellid = 0;
				spreadsheet.createRow(rowid++).createCell(cellid).setCellValue("Deleted Records");
				row = spreadsheet.createRow(rowid++);
				for(String columnName : firstKeySet) {
					Cell cell = row.createCell(cellid++);
			        cell.setCellValue(columnName);
				}
				for(Entry<String, Map<String, String>> deletedRecordsMapEntry : deletedRecordsMap.entrySet()) {
					row = spreadsheet.createRow(rowid++);
					cellid = 0;
					for(Entry<String, String> deletedRecordsMapCellEntry : deletedRecordsMapEntry.getValue().entrySet()) {
						if(firstKeySet.equals(deletedRecordsMapEntry.getValue().keySet())) {
							Cell cell = row.createCell(cellid++);
					        cell.setCellValue(deletedRecordsMapCellEntry.getValue());
						}
					}
				}
			}
			try(FileOutputStream out = new FileOutputStream(new File("CompareResult.xlsx"))) {
		    	workbook.write(out);
			} catch (IOException e) {
				e.printStackTrace();
			}
	    
	}

}
