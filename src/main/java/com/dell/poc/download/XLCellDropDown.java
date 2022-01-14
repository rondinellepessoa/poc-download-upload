package com.dell.poc.download;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;

public class XLCellDropDown {
	public static void main(String args[]) throws FileNotFoundException {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Data Validation");
		CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
		DVConstraint dvConstraint = DVConstraint
				.createExplicitListConstraint(new String[] { "10", "20", "30" });
		DataValidation dataValidation = new HSSFDataValidation(addressList,
				dvConstraint);
		dataValidation.setSuppressDropDownArrow(false);
		sheet.addValidationData(dataValidation);
		FileOutputStream fileOut = new FileOutputStream(
				"XLCellDropDown.xls");
		try {
			workbook.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}

