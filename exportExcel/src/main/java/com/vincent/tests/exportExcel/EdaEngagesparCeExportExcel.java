package com.vincent.tests.exportExcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * classe créant le fichier excel de l'écran EDA Engagés par CE
 * 
 * @author vincent.salomon
 *
 */
public class EdaEngagesparCeExportExcel {

	/**
	 * méthode créant le fichier excel de l'écran EDA Engagés par CE
	 */
	public void ExportExcel() {
		// numéros des lignes et colonnes
		short rownum = 0, colnum = 0, initialrownum;
		// new workbook
		Workbook wb = new XSSFWorkbook();
		// new sheet avec nom passé à l'utilitaire safeName
		String j = "IJ";
		Sheet sheet = wb.createSheet(WorkbookUtil.createSafeSheetName(j));

		// utilitaire d'aide à la création de cellules
		CreationHelper createHelper = wb.getCreationHelper();

		// largeur des colonnes
		sheet.setColumnWidth(0, 100 * 48);
		sheet.setColumnWidth(1, 100 * 48);
		sheet.setColumnWidth(2, 100 * 48);
		sheet.setColumnWidth(3, 100 * 48);
		sheet.setColumnWidth(4, 200 * 48);
		sheet.setColumnWidth(5, 200 * 48);
		sheet.setColumnWidth(6, 200 * 48);

		// titre
		Row rowTitre = sheet.createRow(rownum);
		CellStyle titreStyle = wb.createCellStyle();
		Font titreFont = wb.createFont();
		titreFont.setFontHeightInPoints((short) 12);
		titreFont.setFontName("Verdana");
		titreFont.setColor(IndexedColors.WHITE.getIndex());
		titreStyle.setFont(titreFont);
		titreStyle.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
		titreStyle.setFillPattern( CellStyle.SOLID_FOREGROUND);
		titreStyle.setAlignment(CellStyle.ALIGN_CENTER);
		// LocalDateTime.now().plusDays(1);
		String titre = "EDA engagés le " + LocalDateTime.now().getDayOfMonth() + "/"
				+ LocalDateTime.now().getMonthValue() + "/" + LocalDateTime.now().getYear();
		createCell(createHelper, rowTitre, colnum, titre, titreStyle);
		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, colnum, colnum + 6));
		rownum++;
		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, colnum, colnum + 6));
		rownum++;

		// template tableau
		// style
		CellStyle tabHautCeStyle = wb.createCellStyle();
		CellStyle tabHautStyle = wb.createCellStyle();
		CellStyle tabHautDroitStyle = wb.createCellStyle();
		CellStyle tabCeStyle = wb.createCellStyle();
		CellStyle tabDataStyle = wb.createCellStyle();
		CellStyle tabDataDroitStyle = wb.createCellStyle();
		CellStyle tabBasCEStyle = wb.createCellStyle();
		CellStyle tabBasDataStyle = wb.createCellStyle();
		CellStyle tabBasDroitStyle = wb.createCellStyle();

		Font tabCeFont = wb.createFont();
		Font tabThFont = wb.createFont();
		Font tabDataFont = wb.createFont();

		tabCeFont.setFontHeightInPoints((short) 12);
		tabCeFont.setFontName("Verdana");
		tabCeFont.setColor(IndexedColors.BLACK.getIndex());

		tabThFont.setFontHeightInPoints((short) 12);
		tabThFont.setFontName("Verdana");
		tabThFont.setColor(IndexedColors.BLACK.getIndex());
		tabThFont.setBold(true);

		tabDataFont.setFontHeightInPoints((short) 12);
		tabDataFont.setFontName("Verdana");
		tabDataFont.setColor(IndexedColors.BLACK.getIndex());

		tabHautCeStyle.setFont(tabCeFont);
		tabHautCeStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		tabHautCeStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabHautCeStyle.setAlignment(CellStyle.ALIGN_CENTER);
		tabHautCeStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		tabHautCeStyle.setBorderLeft(CellStyle.BORDER_THIN);
		tabHautCeStyle.setLeftBorderColor(IndexedColors.BROWN.getIndex());
		tabHautCeStyle.setBorderTop(CellStyle.BORDER_THIN);
		tabHautCeStyle.setTopBorderColor(IndexedColors.BROWN.getIndex());

		tabHautStyle.setFont(tabThFont);
		tabHautStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		tabHautStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabHautStyle.setAlignment(CellStyle.ALIGN_CENTER);
		tabHautStyle.setBorderTop(CellStyle.BORDER_THIN);
		tabHautStyle.setTopBorderColor(IndexedColors.BROWN.getIndex());

		tabHautDroitStyle.setFont(tabThFont);
		tabHautDroitStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		tabHautDroitStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabHautDroitStyle.setAlignment(CellStyle.ALIGN_CENTER);
		tabHautDroitStyle.setBorderRight(CellStyle.BORDER_THIN);
		tabHautDroitStyle.setRightBorderColor(IndexedColors.BROWN.getIndex());
		tabHautDroitStyle.setBorderTop(CellStyle.BORDER_THIN);
		tabHautDroitStyle.setTopBorderColor(IndexedColors.BROWN.getIndex());

		tabCeStyle.setBorderLeft(CellStyle.BORDER_THIN);
		tabCeStyle.setLeftBorderColor(IndexedColors.BROWN.getIndex());

		tabDataStyle.setFont(tabDataFont);
		tabDataStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		tabDataStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabDataStyle.setAlignment(CellStyle.ALIGN_CENTER);

		tabDataDroitStyle.setFont(tabDataFont);
		tabDataDroitStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		tabDataDroitStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabDataDroitStyle.setAlignment(CellStyle.ALIGN_CENTER);
		tabDataDroitStyle.setBorderRight(CellStyle.BORDER_THIN);
		tabDataDroitStyle.setRightBorderColor(IndexedColors.BROWN.getIndex());

		tabBasCEStyle.setBorderLeft(CellStyle.BORDER_THIN);
		tabBasCEStyle.setLeftBorderColor(IndexedColors.BROWN.getIndex());
		tabBasCEStyle.setBorderBottom(CellStyle.BORDER_THIN);
		tabBasCEStyle.setBottomBorderColor(IndexedColors.BROWN.getIndex());

		tabBasDataStyle.setFont(tabDataFont);
		tabBasDataStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		tabBasDataStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabBasDataStyle.setAlignment(CellStyle.ALIGN_CENTER);
		tabBasDataStyle.setBorderBottom(CellStyle.BORDER_THIN);
		tabBasDataStyle.setBottomBorderColor(IndexedColors.BROWN.getIndex());

		tabBasDroitStyle.setFont(tabDataFont);
		tabBasDroitStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
		tabBasDroitStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		tabBasDroitStyle.setAlignment(CellStyle.ALIGN_CENTER);
		tabBasDroitStyle.setBorderBottom(CellStyle.BORDER_THIN);
		tabBasDroitStyle.setBottomBorderColor(IndexedColors.BROWN.getIndex());
		tabBasDroitStyle.setBorderRight(CellStyle.BORDER_THIN);
		tabBasDroitStyle.setRightBorderColor(IndexedColors.BROWN.getIndex());

		// data
		initialrownum = rownum;
		Row rowTab = sheet.createRow(rownum++);

		String CE = "CE";
		String thEda = "EDA";
		String thDmo = "DMO";
		String thPuissance = "Puissance";
		String thRo = "RO";
		String thTelPrinc = "N° téléphone principal";
		String thTelSec = "N° téléphone secours";

		createCell(createHelper, rowTab, colnum++, CE, tabHautCeStyle);
		createCell(createHelper, rowTab, colnum++, thEda, tabHautStyle);
		createCell(createHelper, rowTab, colnum++, thDmo, tabHautStyle);
		createCell(createHelper, rowTab, colnum++, thPuissance, tabHautStyle);
		createCell(createHelper, rowTab, colnum++, thRo, tabHautStyle);
		createCell(createHelper, rowTab, colnum++, thTelPrinc, tabHautStyle);
		createCell(createHelper, rowTab, colnum++, thTelSec, tabHautDroitStyle);
		colnum = 0;

		String DataEda = "AAA";
		Double DataDmo = 9.0;
		Double DataPuissance = 1783.0;
		String DataRo = "RO_1";
		String DataTelPrinc = "08 00 00 00 00";
		String DataTelSec = "09 00 00 00 00";

		rowTab = sheet.createRow(rownum++);
		createCell(createHelper, rowTab, colnum++, "", tabCeStyle);
		createCell(createHelper, rowTab, colnum++, DataEda, tabDataStyle);
		createCell(createHelper, rowTab, colnum++, DataDmo, tabDataStyle);
		createCell(createHelper, rowTab, colnum++, DataPuissance, tabDataStyle);
		createCell(createHelper, rowTab, colnum++, DataRo, tabDataStyle);
		createCell(createHelper, rowTab, colnum++, DataTelPrinc, tabDataStyle);
		createCell(createHelper, rowTab, colnum++, DataTelSec, tabDataDroitStyle);
		colnum = 0;

		rowTab = sheet.createRow(rownum);
		createCell(createHelper, rowTab, colnum++, "", tabBasCEStyle);
		createCell(createHelper, rowTab, colnum++, DataEda, tabBasDataStyle);
		createCell(createHelper, rowTab, colnum++, DataDmo, tabBasDataStyle);
		createCell(createHelper, rowTab, colnum++, DataPuissance, tabBasDataStyle);
		createCell(createHelper, rowTab, colnum++, DataRo, tabBasDataStyle);
		createCell(createHelper, rowTab, colnum++, DataTelPrinc, tabBasDataStyle);
		createCell(createHelper, rowTab, colnum++, DataTelSec, tabBasDroitStyle);
		colnum = 0;

		sheet.addMergedRegion(new CellRangeAddress(initialrownum, rownum, colnum, colnum));
		rownum++;
		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, colnum, colnum + 6));
		rownum++;
		initialrownum = rownum;

		// enregisterment du workbook
		try {
			FileOutputStream fileOut = new FileOutputStream("EDA_engages_par_CE.xlsx");
			wb.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	/**
	 * méthode créant une cellule avec attribut d'alignement.
	 *
	 * @param createHelper
	 *            utilitaire d'aide à la création de cellules
	 * @param row
	 *            la ligne de la cellule
	 * @param column
	 *            la colonne de la cellule
	 * @param valeur
	 *            la valeur de la cellule
	 * @param cellStyle
	 *            le style à appliquer à la cellule
	 */
	private static void createCell(CreationHelper createHelper, Row row, short column, String valeur,
			CellStyle cellStyle) {
		Cell cell = row.createCell(column);
		cell.setCellValue(createHelper.createRichTextString(valeur));
		cell.setCellStyle(cellStyle);
	}
	private static void createCell(CreationHelper createHelper, Row row, short column, Double valeur,
			CellStyle cellStyle) {
		Cell cell = row.createCell(column);
		cell.setCellValue(valeur);
		cell.setCellStyle(cellStyle);
	}

}
