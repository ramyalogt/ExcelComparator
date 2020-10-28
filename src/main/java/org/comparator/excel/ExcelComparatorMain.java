package org.comparator.excel;

import java.io.IOException;
import org.comparator.excel.processor.FileParser;

public class ExcelComparatorMain {

	public static void main(String[] args) throws IOException {
		FileParser fileParser = new FileParser(); 
		fileParser.compareFiles();		 
	}
}