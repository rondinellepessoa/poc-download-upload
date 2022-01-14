package com.dell.poc.download;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dell.poc.entity.Produto;

public class ProdutoDownload {

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<Produto> listaProdutos;
	private List<String> listaHeaders;

	public ProdutoDownload (List<Produto> listaProdutos, List<String> hearders) {
        this.listaProdutos = listaProdutos;
        this.listaHeaders = hearders;
        workbook = new XSSFWorkbook();
    }

	private void writeHeaderLine() {
		sheet = workbook.createSheet("Produtos");

		Row row = sheet.createRow(0);

		CellStyle style = workbook.createCellStyle();
		
		CellStyle staticColumnStyle = workbook.createCellStyle();
        staticColumnStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        staticColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(16);
		style.setFont(font);
		staticColumnStyle.setFont(font);
		
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		  // and solid fill pattern produces solid grey cell fill
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		int columnCount = 0;
		for (int i = 0; i < listaHeaders.size(); i++) {
			if(i>1) {
				createCell(row, columnCount, listaHeaders.get(i), staticColumnStyle);

			}else {
				createCell(row, columnCount, listaHeaders.get(i), style);
				
			}
			columnCount++;
		}
	}

	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Long) {
			cell.setCellValue((Long) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Double) {
			cell.setCellValue((Double) value);
		} else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}

	private void writeDataLines() {
		int rowCount = 1;

		CellStyle style = workbook.createCellStyle();
		
		CellStyle staticColumnStyle = workbook.createCellStyle();
        staticColumnStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        staticColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
        CellStyle staticColumnStyle2 = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);

		for (Produto produto : listaProdutos) {
			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;

			createCell(row, columnCount++, produto.getId(), style);			
			createCell(row, columnCount++, produto.getNome(), style);
			createCell(row, columnCount++, produto.getPreco(), staticColumnStyle2);
			createCell(row, columnCount++, produto.getQtde(), staticColumnStyle);
			createCell(row, columnCount++, "333", staticColumnStyle2);
			createCell(row, columnCount++, "555", staticColumnStyle);
			
			//Drop Down list
//			CellRangeAddressList addressList = new CellRangeAddressList(1,5,5,5);
//			DataValidationHelper dvHelper = sheet.getDataValidationHelper();
//			DataValidationConstraint dvConstraint = dvHelper.createExplicitListConstraint(new String[] { "10", "20", "30" });
//			DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
//			validation.setSuppressDropDownArrow(true);
//			sheet.addValidationData(validation);
		}
		
		int colIndex = 2;
		Iterator rowIter = sheet.iterator();
		
        while(rowIter.hasNext()){
        	
            XSSFRow row=(XSSFRow)rowIter.next();
            XSSFCell cell=row.getCell(colIndex);
            row.removeCell(cell);
            
            //sheet.shiftColumns(0,13,1);
             
//    		for (int i = colIndex; i < array.length; i++) {
//    			
//    		}
//            XSSFCell copyCell=row.getCell(colIndex+1);
//            copyCell.setCellStyle(cell.getCellStyle());
//            
//            switch (copyCell.getCellType()) {
//            case STRING:
//            	createCell(row, 2, copyCell.getRichStringCellValue().toString(), style);
//            	break;
//            case BOOLEAN:
//            	createCell(row, 2, copyCell.getRichStringCellValue().toString(), style);
//            	break;
//            case NUMERIC:
//            	createCell(row, 2, copyCell.getNumericCellValue(), style);
//            	break;
//
//            default:
//                throw new RuntimeException("Unexpected cell type (" + copyCell + ")");
//            }
//          
//            colIndex++;
            
        }
        
        List<Integer> listIndex = Arrays.asList(3,4);
        int lastIndex = sheet.getRow(0).getLastCellNum();
        
        listIndex.forEach(index -> {
        	sheet.shiftColumns(index,lastIndex,-1);
        	
        });
        //sheet.shiftColumns(4,4,-1);
        
	}

	public void export(HttpServletResponse response) throws IOException {
		writeHeaderLine();
		writeDataLines();
		
		this.buildResponseDetails(response);

		ServletOutputStream outputStream = response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();

		outputStream.close();
	}
	
	private void buildResponseDetails(HttpServletResponse response) {
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=invoices_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);
    } 

}
