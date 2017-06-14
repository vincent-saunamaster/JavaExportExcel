package com.vincent.tests.exportExcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestsPrealables {
	public void ExportExcel() {
		// new workbook
		Workbook wb = new XSSFWorkbook();
		// new sheet avec nom passé à l'utilitaire safeName
		Sheet sheetIJ = wb.createSheet(WorkbookUtil.createSafeSheetName("IJ"));
		// Sheet sheetJm1 =
		// wb.createSheet(WorkbookUtil.createSafeSheetName("J-1"));

		// utilitaire d'aide à la création de cellules
		CreationHelper createHelper = wb.getCreationHelper();

		//
		// tests préalables
		//

		// nouvelle ligne (commence à 0)
		Row row1er = sheetIJ.createRow((short) 0);

		// nouvelles cellules :
		// soit un double
		row1er.createCell(0).setCellValue(1d);
		// soit un string
		String testString = "test d'un String";
		row1er.createCell(1).setCellValue(createHelper.createRichTextString(testString));

		// application d'un nouveau style pour une cellule et seulement celle
		// -là - cas des dates
		CellStyle styleDateCell = wb.createCellStyle();
		styleDateCell.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
		Cell cell3 = row1er.createCell(2);
		cell3.setCellValue(new Date());
		cell3.setCellStyle(styleDateCell);

		// cellule de type prédéfini
		row1er.createCell(3, CellType.ERROR);

		// initialisation du style d'une cellule qu'on appliquera ensuite
		CellStyle cellStyle = wb.createCellStyle();

		// style du background color
		String j = "IJ";
		if (j.equals("IJ")) {
			cellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		} else {
			cellStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
		}
		cellStyle.setFillPattern((short) CellStyle.SOLID_FOREGROUND);

		// style des fonts
		Font font = wb.createFont();
		font.setFontHeightInPoints((short) 12);
		font.setFontName("Courier New");
		font.setItalic(true);
		font.setStrikeout(true); // barré
		font.setColor(IndexedColors.WHITE.getIndex());
		cellStyle.setFont(font);

		// style des bordures
		cellStyle.setBorderBottom(CellStyle.BORDER_THICK);
		cellStyle.setBottomBorderColor(IndexedColors.BROWN.getIndex());
		cellStyle.setBorderRight(CellStyle.BORDER_THICK);
		cellStyle.setRightBorderColor(IndexedColors.BROWN.getIndex());

		// création d'une cellule avec parametres et valeur dans cellules
		// mergées
		createCell(createHelper, row1er, (short) 4, "valeur String de la cellule", cellStyle);
		// création d'une cellule avec parametres sans valeur dans cellules
		// mergées
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		createCell(createHelper, row1er, (short) 5, null, cellStyle);

		// merge de cellule crée et non crées
		sheetIJ.addMergedRegion(new CellRangeAddress(0, // first row (0-based)
				0, // last row (0-based)
				4, // first column (0-based)
				12 // last column (0-based)
		));
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
}
