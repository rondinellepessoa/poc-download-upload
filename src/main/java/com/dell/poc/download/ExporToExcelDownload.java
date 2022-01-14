package com.dell.poc.download;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dell.poc.entity.Produto;

public class ExporToExcelDownload {

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<T> list;
	private List<String> listaHeaders;

	public ExporToExcelDownload (List<T> list, List<String> hearders) {
        this.list= list;
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
		
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);

		for (T t : list) {
			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;

			createCell(row, columnCount++, t, style);
		}
	}
	
	private void prepareEntityWriteDataLines(Row row, int columnCount, Object t) {
		if(t instanceof Produto) {
			Produto produto = new Produto();
//			produto.setId((Integer) t.);
//			createCell(row, columnCount++, t.getId(), style);			
//			createCell(row, columnCount++, produto.getNome(), style);
//			createCell(row, columnCount++, produto.getPreco(), staticColumnStyle);
//			createCell(row, columnCount++, produto.getQtde(), staticColumnStyle);
//			createCell(row, columnCount++, produto.getCategoria(), staticColumnStyle);
		}
	}

}
