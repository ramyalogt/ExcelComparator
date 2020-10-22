package org.comparator.excel.processor;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Workbook;

public class SheetParser {

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void compareSheet(Workbook oldWorkbook, Workbook newWorkbook, Map sheet) {
		String givenSheetName = ((String) sheet.get("sheetName")).trim();
		String oldWorkbookSheetName = "", newWorkbookSheetName = "";
		if(givenSheetName != null) {
			oldWorkbookSheetName = oldWorkbook.getSheet(givenSheetName).getSheetName().trim();
			newWorkbookSheetName = newWorkbook.getSheet(givenSheetName).getSheetName().trim();
		}
		String sheetName = oldWorkbookSheetName.equals(newWorkbookSheetName) ? oldWorkbookSheetName : "";
		ArrayList<String> givenUniqueKeyColumns = (ArrayList<String>) sheet.get("uniqueKeyColumns");
		LinkedList<String> uniqueKeyColumns = givenUniqueKeyColumns.stream().collect(Collectors.toCollection(LinkedList::new));
		RowParser rowParser = new RowParser();
		if((sheetName != "" && !uniqueKeyColumns.isEmpty()) || (sheetName != null && uniqueKeyColumns != null)) {
			rowParser.compareRows(oldWorkbook, newWorkbook, sheetName, uniqueKeyColumns);
        } else if(sheetName == "" || (sheetName == null && uniqueKeyColumns != null)) 
        	System.out.println("Given Sheet "+givenSheetName+" is not present in the Workbook. Please check and provide valid Sheet name");
        else System.out.println("Please provide Unique Key Columns");
	}
}