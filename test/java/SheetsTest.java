package com.ckkeith.sheetsutility;

import java.util.List;

import junit.framework.TestCase;

public class SheetsTest extends TestCase {

	public SheetsTest(String testName) {
		super(testName);
	}

	public void testRead() {
		List<List<Object>> values = SheetsQuickstart.getRange("1vo9OiDWyDvUGhjaXLUrnr9FkzO2NdUgN-AVe_CONU6U",
				"Sheet1!A1:C15");
		SheetsQuickstart.printData(values);
	}
}
