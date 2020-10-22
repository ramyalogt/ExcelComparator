package org.comparator.excel.processor;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkBookParser {
	
	public XSSFWorkbook oldWorkbook = null;
	public XSSFWorkbook newWorkbook = null;
	public List<String> listOfColumnName = new ArrayList<>();
	public String oldWorkbookColumnName = null;
	public String newWorkbookColumnName = null;
	public String oldWorkbookCellData = null;
	public String newWorkbookCellData = null;
	public String oldWorkbookUniqueKey = "";
	public String newWorkbookUniqueKey = "";
	boolean isPrint = false;
	public Map<String, Map<String, String>> modifiedRecords = new HashMap<>();
	public Map<String, Map<String, String>> deletedRecords = new HashMap<>();
	public Map<String, String> deletedRows= new LinkedHashMap<>();

	@SuppressWarnings({ "rawtypes" })
	public void compareWorkbook(Workbook oldWorkbook, Workbook newWorkbook, List<Map> sheets) {
		this.oldWorkbook = (XSSFWorkbook) oldWorkbook;
		this.newWorkbook = (XSSFWorkbook) newWorkbook;
		SheetParser sheetParser = new SheetParser();
		for (Map sheet : sheets) {
			sheetParser.compareSheet(oldWorkbook, newWorkbook, sheet);
		}
	}
}