package main.java.com.pji.excel.controller;

import java.awt.PageAttributes.MediaType;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import javax.ws.rs.core.Response;

import org.jboss.resteasy.client.ClientRequest;
import org.jboss.resteasy.client.ClientRequestFactory;
import org.jboss.resteasy.client.ClientResponse;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import main.java.com.pji.excel.en.HeaderJson;
import main.java.com.pji.excel.model.Body;
import main.java.com.pji.excel.model.HeaderFooter;
import main.java.com.pji.excel.serviceImpl.BodyImpl;
import main.java.com.pji.excel.serviceImpl.HeaderFooterImpl;
import main.java.com.pji.excel.serviceImpl.ToolImpl;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @category document tool
 * @version 1.0
 * @param <T>
 */
public class DocumentHandler {

	/**
	 * @author arief.pramudyo
	 * @param fileName
	 * @return WritableWorkbook
	 * @since 12/21/2015
	 */
	public WritableWorkbook createdWorkBook(String fileName) {
		WritableWorkbook workbook = null;
		if (fileName != null) {
			try {
				File file = new File(fileName);
				WorkbookSettings wbSettings = new WorkbookSettings();

				wbSettings.setLocale(new Locale("en", "EN"));

				workbook = Workbook.createWorkbook(file, wbSettings);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return workbook;
	}

	/**
	 * @author arief.pramudyo
	 * @param writableWorkbook
	 * @param nameSheet
	 * @return WritableSheet
	 * @since 12/21/2015
	 */
	public WritableSheet createSheet(WritableWorkbook writableWorkbook,
			String nameSheet) {
		if (nameSheet == null)
			writableWorkbook.createSheet("sheet1", 0);
		else
			writableWorkbook.createSheet(nameSheet, 0);

		WritableSheet excelSheet = writableWorkbook.getSheet(0);
		return excelSheet;
	}

	private HeaderFooter addHeaderFooter(String header,
			String positionHeader, WritableCellFormat writableCellFormatHeader,
			String footer, String positionFooter,
			WritableCellFormat writableCellFormatFooter) {
		HeaderFooter headerFooter = new HeaderFooter();
		if (header != null) {
			headerFooter.setHeader(header);
		}
		if (positionHeader != null) {
			headerFooter.setPositionHeader(positionHeader);
		}
		if (writableCellFormatHeader != null) {
			headerFooter.setWritableCellFormatHeader(writableCellFormatHeader);
		}
		if (footer != null) {
			headerFooter.setFooter(footer);
		}
		if (positionFooter != null) {
			headerFooter.setPositionFooter(positionFooter);
		}
		if (writableCellFormatFooter != null) {
			headerFooter.setWritableCellFormatFooter(writableCellFormatFooter);
		}
		return headerFooter;
	}

	public HeaderFooter addHeader(String headerValue, String positionHeader,
			WritableCellFormat writableCellFormatHeader) {
		return addHeaderFooter(headerValue, positionHeader,
				writableCellFormatHeader, null, null, null);
	}

	public HeaderFooter addFooter(String footerValue, String positionFooter,
			WritableCellFormat writableCellFormatFooter) {
		return addHeaderFooter(footerValue, positionFooter,
				writableCellFormatFooter, null, null, null);
	}

	public Integer writeHeaderFooter(WritableSheet sheet,
			List<HeaderFooter> modelHeaderFooter, Integer rightPosition) {
		Integer sizeHeaderFooter = null;
		try {
			HeaderFooterImpl headerFootersImpl = new HeaderFooterImpl();
			if (modelHeaderFooter != null) {
				sizeHeaderFooter = headerFootersImpl.addHeaderMatriks(sheet,
						modelHeaderFooter, rightPosition);
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return sizeHeaderFooter;
	}

	/**
	 * @author arief.pramudyo
	 * @param excelSheet
	 * @param writableCellFormat
	 * @param label
	 * @param labelHeader
	 * @param row
	 * @return WritableSheet
	 * @since 12/21/2015
	 */
	public WritableSheet writeExcel(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize) {
		try {
			BodyImpl bodyImpl = new BodyImpl();
			if (listLabel != null) {
				bodyImpl.addBody(sheet, row, listLabel,
						writableCellFormatLabel, listCaption,
						writableCellFormatCaption, headerRowSize);
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return sheet;
	}

	/**
	 * @author arief.pramudyo
	 * @param writableWorkbook
	 * @since 12/21/2015
	 */
	public void executeWriteExcel(WritableWorkbook writableWorkbook) {
		try {
			writableWorkbook.write();
			writableWorkbook.close();
		} catch (IOException | WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 
	 * @author arief.pramudyo
	 * @param nameSheet
	 * @param fileName
	 * @param bodyModel
	 * @param row
	 * @param listLabel
	 * @param writableCellFormatLabel
	 * @param listCaption
	 * @param writableCellFormatCaption
	 * @param headerRowSize
	 * @since 12/21/2015
	 */
	public void simpleToExcelWithHeaderFooter(String nameSheet, String fileName,
			Body bodyModel, Boolean row, List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption,
			List<HeaderFooter> model) {
		WritableWorkbook workbook = createdWorkBook(fileName);
		WritableSheet sheet = createSheet(workbook, nameSheet);
		Integer headerFooterSizeRow = null;
		if (listCaption != null)
			headerFooterSizeRow=writeHeaderFooter(sheet, model, listCaption.size()-1);
		writeExcel(sheet, row, listLabel, writableCellFormatLabel, listCaption,
				writableCellFormatCaption, headerFooterSizeRow);
		executeWriteExcel(workbook);
	}

	public void simpleToExcel(String nameSheet, String fileName,
			Body bodyModel, Boolean row, List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize) {
		WritableWorkbook workbook = createdWorkBook(fileName);
		WritableSheet sheet = createSheet(workbook, nameSheet);
		writeExcel(sheet, row, listLabel, writableCellFormatLabel, listCaption,
				writableCellFormatCaption, headerRowSize);
		executeWriteExcel(workbook);
	}

	/**
	 * 
	 * @author arief.pramudyo
	 * @param nameSheet
	 * @param fileName
	 * @param bodyModel
	 * @param row
	 * @param listLabel
	 * @param writableCellFormatLabel
	 * @param listCaption
	 * @param writableCellFormatCaption
	 * @since 12/21/2015
	 */
	public void simpleToExcel(String nameSheet, String fileName,
			Body bodyModel, Boolean row, List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption) {

		simpleToExcel(nameSheet, fileName, bodyModel, row, listLabel,
				writableCellFormatLabel, listCaption,
				writableCellFormatCaption, null);
	}

	/**
	 * 
	 * @author arief.pramudyo
	 * @param nameSheet
	 * @param fileName
	 * @param bodyModel
	 * @param row
	 * @param listLabel
	 * @param writableCellFormatLabel
	 * @param listCaption
	 * @since 12/21/2015
	 */
	public void simpleToExcel(String nameSheet, String fileName,
			Body bodyModel, Boolean row, List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel, List<String> listCaption) {
		simpleToExcel(nameSheet, fileName, bodyModel, row, listLabel,
				writableCellFormatLabel, listCaption, null, null);
	}

	/**
	 * 
	 * @author arief.pramudyo
	 * @param nameSheet
	 * @param fileName
	 * @param bodyModel
	 * @param row
	 * @param listLabel
	 * @param writableCellFormatLabel
	 * @since 12/21/2015
	 */
	public void simpleToExcel(String nameSheet, String fileName,
			Body bodyModel, Boolean row, List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel) {
		simpleToExcel(nameSheet, fileName, bodyModel, row, listLabel,
				writableCellFormatLabel, null, null, null);
	}

	/**
	 * 
	 * @author arief.pramudyo
	 * @param nameSheet
	 * @param fileName
	 * @param bodyModel
	 * @param row
	 * @param listLabel
	 * @since 12/21/2015
	 */
	public void simpleToExcel(String nameSheet, String fileName,
			Body bodyModel, Boolean row, List<List<String>> listLabel) {
		simpleToExcel(nameSheet, fileName, bodyModel, row, listLabel, null,
				null, null, null);
	}

	/**
	 * 
	 * @author arief.pramudyo
	 * @param nameSheet
	 * @param fileName
	 * @param bodyModel
	 * @param listLabel
	 * @since 12/21/2015
	 */
	public void simpleToExcel(String nameSheet, String fileName,
			Body bodyModel, List<List<String>> listLabel) {
		simpleToExcel(nameSheet, fileName, bodyModel, true, listLabel, null,
				null, null, null);
	}

	/**
	 * @author arief.pramudyo
	 * @param cvsName
	 * @param excelName
	 * @since 12/21/2015
	 */
	// execute Import Excell to Cvs
	// format cvs name === drive:/nameFolder/nameFile.cvs
	// format excel name === drive:/nameFolder/excelFile.xls
	public void excelToCvs(String cvsName, String excelName) {
		try {
			// File to store data in form of CSV
			File f = new File(cvsName);

			@SuppressWarnings("resource")
			OutputStream os = (OutputStream) new FileOutputStream(f);
			String encoding = "UTF8";
			OutputStreamWriter osw = new OutputStreamWriter(os, encoding);
			BufferedWriter bw = new BufferedWriter(osw);

			// Excel document to be imported
			String filename = excelName;
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			Workbook w = Workbook.getWorkbook(new File(filename), ws);

			// Gets the sheets from workbook
			for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++) {
				Sheet s = w.getSheet(sheet);

				bw.write(s.getName());
				bw.newLine();

				Cell[] row = null;

				// Gets the cells from sheet
				for (int i = 0; i < s.getRows(); i++) {
					row = s.getRow(i);

					if (row.length > 0) {
						bw.write(row[0].getContents());
						for (int j = 1; j < row.length; j++) {
							bw.write(',');
							bw.write(row[j].getContents());
						}
					}
					bw.newLine();
				}
			}
			bw.flush();
			bw.close();
		} catch (Exception e) {
			System.err.println(e.toString());
		}
	}

	/**
	 * @author arief.pramudyo
	 * @param excelName
	 * @return Map<String, List<String>>
	 * @since 12/21/2015
	 */
	// read Excel File More One sheet
	public Map<String, List<String>> readExcelManySheet(String excelName) {
		Map<String, List<String>> mapExcell = new HashMap<String, List<String>>();
		try {
			List<String> listCellString = new ArrayList<String>();
			// Excel document to be imported
			String filename = excelName;
			String valueColumnCell = "";
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			Workbook w = Workbook.getWorkbook(new File(filename), ws);

			// Gets the sheets from workbook
			for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++) {
				Sheet s = w.getSheet(sheet);

				s.getName();

				Cell[] row = null;

				// Gets the cells from sheet
				for (int i = 0; i < s.getRows(); i++) {
					row = s.getRow(i);

					if (row.length > 0) {
						for (int j = 1; j < row.length; j++) {
							valueColumnCell = ',' + row[j].getContents();
						}
						listCellString.add(row[0].getContents()
								+ valueColumnCell);
					}
				}
				mapExcell.put(s.getName(), listCellString);
			}
		} catch (Exception e) {
			System.err.println(e.toString());
		}
		return mapExcell;
	}

	/**
	 * @author arief.pramudyo
	 * @param excelName
	 * @return List<String>
	 * @since 12/21/2015
	 */
	// read Excel File More One sheet
	public List<String> readExcelOneSheet(String excelName) {
		List<String> listCellString = new ArrayList<String>();
		try {
			// Excel document to be imported
			String filename = excelName;
			String valueColumnCell = "";
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			Workbook w = Workbook.getWorkbook(new File(filename), ws);

			// Gets the sheets from workbook
			Sheet s = w.getSheet(0);

			s.getName();

			Cell[] row = null;

			// Gets the cells from sheet
			for (int i = 0; i < s.getRows(); i++) {
				row = s.getRow(i);

				if (row.length > 0) {
					for (int j = 1; j < row.length; j++) {
						valueColumnCell = ',' + row[j].getContents();
					}
					listCellString.add(row[0].getContents() + valueColumnCell);
				}
			}
		} catch (Exception e) {
			System.err.println(e.toString());
		}
		return listCellString;
	}

	/**
	 * @author arief.pramudyo
	 * @param pdfName
	 * @param excelName
	 * @since 12/21/2015
	 */
	// execute Import Excell to PDF
	// format pdf name === drive:/nameFolder/nameFile.pdf
	// format excel name === drive:/nameFolder/excelFile.xls
	public void excelToPdf(String pdfName, String excelName) {
		try {
			// File to store data in form of CSV
			File f = new File(pdfName);
			Document iText_xls_2_pdf = new Document();
			PdfWriter.getInstance(iText_xls_2_pdf,
					(OutputStream) new FileOutputStream(f));
			iText_xls_2_pdf.open();

			PdfPTable my_table;
			PdfPCell table_cell;

			// Excel document to be imported
			String filename = excelName;
			WorkbookSettings ws = new WorkbookSettings();
			ws.setLocale(new Locale("en", "EN"));
			Workbook w = Workbook.getWorkbook(new File(filename), ws);

			// Gets the sheets from workbook
			for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++) {
				Sheet s = w.getSheet(sheet);
				my_table = new PdfPTable(s.getColumns());
				Cell[] row = null;

				// Gets the cells from sheet
				for (int i = 0; i < s.getRows(); i++) {
					row = s.getRow(i);

					if (row.length > 0) {
						table_cell = new PdfPCell(new Phrase(
								row[0].getContents()));
						my_table.addCell(table_cell);
						for (int j = 1; j < row.length; j++) {
							table_cell = new PdfPCell(new Phrase(
									row[0].getContents()));
							my_table.addCell(table_cell);
						}
					}
				}
				iText_xls_2_pdf.add(my_table);
				iText_xls_2_pdf.newPage();
			}
			iText_xls_2_pdf.close();
		} catch (Exception e) {
			System.err.println(e.toString());
		}
	}
	
	/**
	 * @author arief.pramudyo
	 * @param pdfName
	 * @param workbook
	 * @since 12/21/2015
	 */
	// execute Import Workbook to PDF
	// format pdf name === drive:/nameFolder/nameFile.pdf
	public void workbookToPdf(String pdfName, WritableWorkbook workbook) {
		try {
			// File to store data in form of CSV
			File f = new File(pdfName);
			Document iText_xls_2_pdf = new Document();
			PdfWriter.getInstance(iText_xls_2_pdf,
					(OutputStream) new FileOutputStream(f));
			iText_xls_2_pdf.open();

			PdfPTable my_table;
			PdfPCell table_cell;

			// Gets the sheets from workbook
			for (int sheet = 0; sheet < workbook.getNumberOfSheets(); sheet++) {
				Sheet s = workbook.getSheet(sheet);
				my_table = new PdfPTable(s.getColumns());	
				my_table.setWidthPercentage(100f);
				Cell[] row = null;
				Font font = new Font(FontFamily.TIMES_ROMAN, 12);
				// Gets the cells from sheet
				for (int i = 0; i < s.getRows(); i++) {
					row = s.getRow(i);
					if (row.length > 0) {
						table_cell = new PdfPCell(new Phrase(
								row[0].getContents(),font));
						my_table.addCell(table_cell);						
						if(row.length<s.getColumns()){
							for (int j = 1; j < s.getColumns(); j++) {
								table_cell = new PdfPCell(new Phrase(
										"",font));
								my_table.addCell(table_cell);
							}
						}
						for (int j = 1; j < row.length; j++) {
							table_cell = new PdfPCell(new Phrase(
									row[j].getContents(),font));
							my_table.addCell(table_cell);
						}
					}
				}
				iText_xls_2_pdf.add(my_table);
				iText_xls_2_pdf.newPage();
			}
			iText_xls_2_pdf.close();
		} catch (Exception e) {
			System.err.println(e.toString());
		}
	}

	/**
	 * @author arief.pramudyo
	 * @param nameCvsFile
	 * @param nameExcelFile
	 * @since 12/21/2015
	 */
	// execute Import Cvs To Excell
	public void cvsToExcell(String nameCvsFile, String nameExcelFile,
			Boolean fileCvsOnResouceFolde) {
		BufferedReader br = null;
		Integer i = 0;
		List<List<String>> bodyList = new ArrayList<List<String>>();
		List<String> body = new ArrayList<String>();
		WritableWorkbook workbook = createdWorkBook(nameExcelFile);
		WritableSheet sheet = createSheet(workbook, null);
		try {

			String sCurrentLine;

			// example /file/textCvs.txt
			if (fileCvsOnResouceFolde == true)
				br = new BufferedReader(new FileReader(getClass()
						.getClassLoader().getResource(nameCvsFile).getFile()));
			else
				br = new BufferedReader(new FileReader(nameCvsFile));

			while ((sCurrentLine = br.readLine()) != null) {
				for (int j = 0; j < sCurrentLine.split(",").length; j++) {
					body.add(sCurrentLine.split(",")[j]);
				}
				bodyList.add(body);
				i++;
			}
			writeExcel(sheet, true, bodyList, null, null, null, null);
			executeWriteExcel(workbook);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (br != null)
					br.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
	}

	/**
	 * @author arief.pramudyo
	 * @param nameCvsFile
	 * @return Map<Integer, String>
	 * @since 12/21/2015
	 */
	// execute Import Cvs
	public Map<Integer, String> readCvs(String nameCvsFile,
			Boolean fileInResourceFolder) {
		BufferedReader br = null;
		Integer i = 0;
		Map<Integer, String> mapCvs = new HashMap<Integer, String>();
		try {

			String sCurrentLine;
			// example /file/textCvs.txt
			if (fileInResourceFolder == true)
				br = new BufferedReader(new FileReader(getClass()
						.getClassLoader().getResource(nameCvsFile).getFile()));
			else
				br = new BufferedReader(new FileReader(nameCvsFile));

			while ((sCurrentLine = br.readLine()) != null) {
				mapCvs.put(i, sCurrentLine);
				i++;
			}

		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (br != null)
					br.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
		return mapCvs;
	}


}
