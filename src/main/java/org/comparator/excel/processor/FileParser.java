package org.comparator.excel.processor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.esotericsoftware.yamlbeans.YamlException;
import com.esotericsoftware.yamlbeans.YamlReader;

@SuppressWarnings({ "unchecked", "rawtypes" })
public class FileParser {
	
	YamlReader reader = null;
	Map map = null;
	String oldFileName = null;
	String newFileName = null;
	File oldFile = null;
	File newFile = null;
	List<Map> sheets = null;
	
	public FileParser() throws FileNotFoundException, YamlException {
		this.reader = new YamlReader(new FileReader("application.yaml"));
		this.map = (Map) reader.read();
		this.sheets = getSheets();
		getExcelFiles();
	}
	
	public void compareFiles() throws IOException {
		WorkBookParser workBookParser = new WorkBookParser();
		try(
        		FileInputStream oldFileInputStream = new FileInputStream(oldFile);
        		FileInputStream newFileInputStream = new FileInputStream(newFile)
        	) {
        	Workbook oldWorkbook = new XSSFWorkbook(oldFileInputStream);
        	Workbook newWorkbook = new XSSFWorkbook(newFileInputStream);
        	workBookParser.compareWorkbook(oldWorkbook, newWorkbook, sheets);
        } catch (Exception e) {
            e.printStackTrace();
        }
		
	}

	private List<Map> getSheets() {
		List<Map> sheets = (List<Map>) this.map.get("sheets");
		return sheets;
	}

	private void getExcelFiles() {
		this.map.forEach((key,value) -> {
			if(value.getClass().equals(String.class)) {
            	if(key.equals("oldFileName")) this.oldFile = new File((String)value);
            	if(key.equals("newFileName")) this.newFile = new File((String)value);
			}
		});
	}
}