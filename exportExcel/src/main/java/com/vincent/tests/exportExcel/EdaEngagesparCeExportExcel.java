package com.vincent.tests.exportExcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;

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
	public void exportExcel() {
		// numéros des lignes et colonnes
		short rownum = 0, colnum = 0, initialcolnum = 0, initialrownum;
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
		titreStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
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
		CellStyle tabBasStyle = wb.createCellStyle();

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

		tabBasStyle.setBorderTop(CellStyle.BORDER_THIN);
		tabBasStyle.setTopBorderColor(IndexedColors.BROWN.getIndex());

		// data
		//
		// à reproduire
		// $scope.listeJbyCE = _.groupBy(_.sortBy(_.sortBy(response.data,
		// 'libEda'), 'libCE'), 'libCE');

		// création données
		class ObjetTest {
			String ce;
			String eda;
			String dmo;
			String puissance;
			String ro;
			String telPrinc;
			String telSecours;

			public String getCe() {
				return ce;
			}

			public void setCe(String ce) {
				this.ce = ce;
			}

			public String getEda() {
				return eda;
			}

			public void setEda(String eda) {
				this.eda = eda;
			}

			public String getDmo() {
				return dmo;
			}

			public void setDmo(String dmo) {
				this.dmo = dmo;
			}

			public String getPuissance() {
				return puissance;
			}

			public void setPuissance(String puissance) {
				this.puissance = puissance;
			}

			public String getRo() {
				return ro;
			}

			public void setRo(String ro) {
				this.ro = ro;
			}

			public String getTelPrinc() {
				return telPrinc;
			}

			public void setTelPrinc(String telPrinc) {
				this.telPrinc = telPrinc;
			}

			public String getTelSecours() {
				return telSecours;
			}

			public void setTelSecours(String telSecours) {
				this.telSecours = telSecours;
			}

		}
		ObjetTest a = new ObjetTest();
		a.setCe("ceA");
		a.setEda("edad");
		a.setDmo("dmo1");
		a.setPuissance("puissance1");
		a.setRo("ro1");
		a.setTelPrinc("telPrinc1");
		a.setTelSecours("telSecours1");
		ObjetTest b = new ObjetTest();
		b.setCe("ceA");
		b.setEda("edaf");
		b.setDmo("dmo2");
		b.setPuissance("puissance2");
		b.setRo("ro2");
		b.setTelPrinc("telPrinc2");
		b.setTelSecours("telSecours2");
		ObjetTest c = new ObjetTest();
		c.setCe("ceB");
		c.setEda("edae");
		c.setDmo("dmo3");
		c.setPuissance("puissance3");
		c.setRo("ro3");
		c.setTelPrinc("telPrinc3");
		c.setTelSecours("telSecours3");

		List<ObjetTest> listTest = new ArrayList<>();
		listTest.add(c);
		listTest.add(b);
		listTest.add(a);

		// sort et groupBy
		listTest.sort(Comparator.comparing(ObjetTest::getEda));
		Map<String, List<ObjetTest>> MapgroupByCeUnSorted = listTest.stream().collect(Collectors.groupingBy(ObjetTest::getCe));
		Map<String, List<ObjetTest>> MapgroupByCeSort = MapgroupByCeUnSorted.entrySet().stream()
				.sorted(Map.Entry.comparingByKey()).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue,
						(oldValue, newvalue) -> oldValue, LinkedHashMap::new));

		// valeurs pour le parcours des tableaux
		initialrownum = rownum;
		Row rowTab;

		// nom des Th
		String CE = "CE";
		String thEda = "EDA";
		String thDmo = "DMO";
		String thPuissance = "Puissance";
		String thRo = "RO";
		String thTelPrinc = "N° téléphone principal";
		String thTelSec = "N° téléphone secours";

		for (Map.Entry<String, List<ObjetTest>> entry : MapgroupByCeSort.entrySet()) {
			rowTab = sheet.createRow(rownum++);
			createCell(createHelper, rowTab, colnum++, entry.getKey(), tabHautCeStyle);
			createCell(createHelper, rowTab, colnum++, thEda, tabHautStyle);
			createCell(createHelper, rowTab, colnum++, thDmo, tabHautStyle);
			createCell(createHelper, rowTab, colnum++, thPuissance, tabHautStyle);
			createCell(createHelper, rowTab, colnum++, thRo, tabHautStyle);
			createCell(createHelper, rowTab, colnum++, thTelPrinc, tabHautStyle);
			createCell(createHelper, rowTab, colnum++, thTelSec, tabHautDroitStyle);
			colnum = initialcolnum;

			for (ObjetTest o : entry.getValue()) {
				rowTab = sheet.createRow(rownum++);
				createCell(createHelper, rowTab, colnum++, "", tabCeStyle);
				createCell(createHelper, rowTab, colnum++, o.getEda(), tabDataStyle);
				createCell(createHelper, rowTab, colnum++, o.getDmo(), tabDataStyle);
				createCell(createHelper, rowTab, colnum++, o.getPuissance(), tabDataStyle);
				createCell(createHelper, rowTab, colnum++, o.getRo(), tabDataStyle);
				createCell(createHelper, rowTab, colnum++, o.getTelPrinc(), tabDataStyle);
				createCell(createHelper, rowTab, colnum++, o.getTelSecours(), tabDataDroitStyle);
				colnum = initialcolnum;
			}
			// merge colone CE
			sheet.addMergedRegion(new CellRangeAddress(initialrownum, rownum - 1, colnum, colnum));

			rowTab = sheet.createRow(rownum++);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			createCell(createHelper, rowTab, colnum++, "", tabBasStyle);
			colnum = initialcolnum;

			sheet.addMergedRegion(new CellRangeAddress(rownum - 1, rownum - 1, colnum, colnum + 6));
			initialrownum = rownum;

		}

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
