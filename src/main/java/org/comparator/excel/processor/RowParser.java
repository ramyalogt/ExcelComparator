package org.comparator.excel.processor;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ss.usermodel.Workbook;

public class RowParser {

	public void compareRows(Workbook oldWorkbook, Workbook newWorkbook, String sheetName,
			LinkedList<String> uniqueKeyColumns) {
		CellParser cellParser = new CellParser(oldWorkbook, newWorkbook, sheetName, uniqueKeyColumns);
		
		//(uniquekey, uniqueKeyvaluesList)  
		Map<String, List<String>> uniqueKeyMap = cellParser.getCommonUniqueKeys();
		
		Map<String, Map<String, Map<String, String>>> modifiedAndDeletedValuesMap = cellParser.getModifiedAndDeletedValues(uniqueKeyMap);
		if(!modifiedAndDeletedValuesMap.isEmpty() && modifiedAndDeletedValuesMap != null) {
			System.out.println("Modified and Deleted Columns(Column {OldValue|NewValue}):");
			int i = 1;
			for(Entry<String, Map<String, Map<String, String>>> modifiedAndDeletedValuesMapEntry : modifiedAndDeletedValuesMap.entrySet()) {
				String lastKey = null;
				if(!modifiedAndDeletedValuesMapEntry.getValue().isEmpty()) {
					for(String key : modifiedAndDeletedValuesMapEntry.getValue().keySet()){
						lastKey = key;
					  }
				}
				System.out.print(i+". ");
				for(Entry<String, Map<String, String>> modifiedColumnsEntry : modifiedAndDeletedValuesMapEntry.getValue().entrySet()) {
					for(Entry<String, String> oldAndNewValuesEntry : modifiedColumnsEntry.getValue().entrySet()) {
						if(lastKey.equals(modifiedColumnsEntry.getKey())) System.out.print(modifiedColumnsEntry.getKey()+"={"+oldAndNewValuesEntry.getKey()+"|"+oldAndNewValuesEntry.getValue()+"}");
						else System.out.print(modifiedColumnsEntry.getKey()+"={"+oldAndNewValuesEntry.getKey()+"|"+oldAndNewValuesEntry.getValue()+"}, ");
						
					}
				} 
				i++;
				System.out.println();
			}
			System.out.println();
		}
		
		
		Map<String, Map<String, String>> insertedRowsMap= cellParser.getInsertedRows(uniqueKeyMap);
		if(!insertedRowsMap.isEmpty() && insertedRowsMap != null) {
			System.out.println("Inserted Rows(Column=Value):");
			int i = 1;
			for(Entry<String, Map<String, String>> insertedRowsMapEntry : insertedRowsMap.entrySet()) {
				System.out.println(i+". "+insertedRowsMapEntry.getValue());
				i++;
			}
			System.out.println();
		}
		
		Map<String, Map<String, String>> deletedRowsMap = cellParser.getDeletedRows(uniqueKeyMap);
		if(!deletedRowsMap.isEmpty() && deletedRowsMap != null) {
			System.out.println("Deleted Rows(Column=Value):");
			int i = 1;
			for(Entry<String, Map<String, String>> deletedRowsMapEntry : deletedRowsMap.entrySet()) {
				System.out.println(i+". "+deletedRowsMapEntry.getValue());
				i++;
			}
			System.out.println();
		}
	}

}
