package org.hectormoraga.apache.demo;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A dirty simple program that reads an Excel file.
 * 
 * @author www.codejava.net
 *
 */
public class SimpleExcelReaderExample {

	public static void main(String[] args) throws IOException {
		String excelFilePath = "AnualIncome.xlsx";

		// genero un doc pdf vacio
		PDDocument document = new PDDocument();
		// doy formato de una pagina tamaño carta
		PDPage page = new PDPage(PDRectangle.LEGAL);
		// le agrego al documento la pagina
		document.addPage(page);
		// calculo largo y alto de la pagina
		final int pageHeight = (int) page.getTrimBox().getHeight();
		final int pageWidth = (int) page.getTrimBox().getWidth();
		// para habilitar el flujo del contenido
		PDPageContentStream contentStream = new PDPageContentStream(document, page);
		//
		contentStream.setStrokingColor(Color.black);
		// ancho de las lineas
		contentStream.setLineWidth(1);

		// flujo de entrada del archivo excel a leer
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		// ???
		Workbook workbook = new XSSFWorkbook(inputStream);
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator(); 

		// leo la primera sabana del archivo excel (suponiendo que no tiene mas)
		Sheet firstSheet = workbook.getSheetAt(0);

		int maxRows = firstSheet.getPhysicalNumberOfRows();
		
		// para averiguar el maximo de columnas debo darme una vuelta completa por todas
		// las columnas para contar
		int maxCols = 0;
		
		for (int rows=0;rows<maxRows;rows++) {
			if (maxCols<firstSheet.getRow(rows).getLastCellNum()) {
				maxCols=firstSheet.getRow(rows).getLastCellNum();
			}
		}
		
		System.out.println("rows:" + maxRows + ", cols:" + maxCols);
		// ???
		int initX = 50;
		int initY = pageHeight - 50;
		int cellWidth = (pageWidth-2*initX) / maxCols;

		for (int i = 0; i < maxRows; i++) {
			XSSFRow row = (XSSFRow) firstSheet.getRow(i);
			// altura en pt
			int cellHeight = (int)row.getHeightInPoints();
			System.out.println("cell height:" + cellHeight+ " width:" + cellWidth);
			
			for (int j = firstSheet.getFirstRowNum(); j < maxCols; j++) {
				contentStream.addRect(initX, initY, cellWidth, -cellHeight);

				Cell cell = row.getCell(j);
				
				if (cell!=null) {
					System.out.print(cell.toString()+" ");

					contentStream.beginText();

					contentStream.newLineAtOffset(initX + 10, initY - cellHeight+4);
					contentStream.setFont(PDType1Font.TIMES_ROMAN, 10);
					contentStream.showText(getCellValue(evaluator, cell));

					contentStream.endText();
				}
				initX += cellWidth;
			}
			initX = 50;
			
			if (initY>cellHeight) {
				initY -= cellHeight;				
			} else {
				contentStream.stroke();
				contentStream.close();
				
				page = new PDPage(PDRectangle.LEGAL);
				document.addPage(page);
				
				initX = 50;
				initY = pageHeight - 50;
				cellWidth = (pageWidth-2*initX) / maxCols;
				
				contentStream = new PDPageContentStream(document, page);
			}
			System.out.println();
		}

		inputStream.close();
		workbook.close();

		contentStream.stroke();
		contentStream.close();

		document.save("output.pdf");
		document.close();
	}

	private static String getCellValue(FormulaEvaluator evaluator, Cell cell) {
		String retorno = "";

		if (cell!=null) {			
			switch (cell.getCellType()) {
			case STRING:
				retorno = cell.getStringCellValue();
				break;
			case BOOLEAN:
				retorno = Boolean.toString(cell.getBooleanCellValue());
				break;
			case NUMERIC:
				retorno = BigInteger.valueOf((long)cell.getNumericCellValue()).toString();
				break;
			case FORMULA:
				switch (evaluator.evaluateFormulaCell(cell)) {
				case BOOLEAN:
					retorno = Boolean.toString(cell.getBooleanCellValue());
					break;
				case NUMERIC:
					retorno = BigInteger.valueOf((long)cell.getNumericCellValue()).toString();
					break;
				case STRING:
					retorno = cell.getStringCellValue();
					break;
				default:
					break;
				}
				break;
			default:
				break;
			}
		}
		
		return retorno;
	}
}