package org.comparator.excel.test;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.comparator.excel.processor.FileParser;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import com.esotericsoftware.yamlbeans.YamlException;

public class ExcelComparatorTest {
	
	public static FileParser fileParser;
	
	@BeforeAll
	public static void init() {
		try {
			fileParser = new FileParser();
		} catch (FileNotFoundException | YamlException e) {
			e.printStackTrace();
		}
	}
	
	
	@Test
	public void getModifiedAndDeletedValuesTest() {
		try {
			fileParser.compareFiles();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
