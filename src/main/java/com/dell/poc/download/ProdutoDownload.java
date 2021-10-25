package com.dell.poc.download;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
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
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(16);
		style.setFont(font);
		
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		  // and solid fill pattern produces solid grey cell fill
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		int columnCount = 0;
		for (int i = 0; i < listaHeaders.size(); i++) {
			createCell(row, columnCount, listaHeaders.get(i), style);
			columnCount++;
		}

//		createCell(row, 0, "Produto ID", style);
//		createCell(row, 1, "Nome", style);
//		createCell(row, 2, "Preço", style);

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
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);

		for (Produto produto : listaProdutos) {
			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;

			createCell(row, columnCount++, produto.getId(), style);
			createCell(row, columnCount++, produto.getNome(), style);
			createCell(row, columnCount++, produto.getPreco(), style);

		}
	}

	public void export(HttpServletResponse response) throws IOException {
		writeHeaderLine();
		writeDataLines();

		ServletOutputStream outputStream = response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();

		outputStream.close();

	}

}
