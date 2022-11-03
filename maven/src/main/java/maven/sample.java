package maven;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class sample {
	public static String defaultfontStyle = "Calibri";
	public static short defaultfontColor = IndexedColors.BLACK.getIndex() ;
	public static void main(String[] args) throws IOException {

		Runtime.getRuntime().exec("cmd /c taskkill /f /im excel.exe");
		// String path = "C:\\WorqForce\\output\\Summarysheet.xlsx";
		// String path = "C:\\WorqForce\\output\\Summarysheet.xlsx";
		String filePath = "C:\\WorqForce\\output\\";
		String fileName = "Summarysheet";
		String extension = ".xlsx";

		String date = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss").format(new Date());

		date = date.replaceAll("/", "_");
		date = date.replaceAll(" ", "_");
		date = date.replaceAll(":", "_");

		String path = filePath + fileName + "_" + date + extension;
		String filePath2 = FilenameUtils.getFullPath(path);
		System.out.println(filePath2);
		XSSFSheet sheet = null;
		XSSFWorkbook workbook = null;
		String Sheetname = "success";
		int rowcount = 1;
		List<String> headers1 = new ArrayList<String>();
		List<String> headers2 = new ArrayList<String>();
		headers1.add("BOT Run Result");
		headers1.add(
				"Variance Analysis:\n 1) negative amount \n 2) +/- 30% variance for EBITDAR, G&A, Sales \n 3) +/- 10% variance for Rent");

		headers2.add("RI#");
		headers2.add("Last Processed date");
		headers2.add("Status");
		headers2.add("Business Exception");
		headers2.add("Sales Exception");
		headers2.add("EBITDAR Exception");
		headers2.add("G&A Exception");
		headers2.add("R&D Exception");
		headers2.add("Sales Threshold Exception");
		headers2.add("EBITDAR Threshold Exception");
		headers2.add("G&A Threshold Exception");
		headers2.add("R&D Threshold Exception");

		List<String> values = new ArrayList<String>();

		values.add(
				"26342,09-29-2022,failure,,,,,,-68% Sales Value is not within the threshold range.,-63100.74,,");
		values.add(
				"263433,09-29-2022,success,,,,,,-68% Sales Value is not within the threshold range.,-63100.74  ,,");
		values.add(
				"2634,09-29-2022,failure,,,,,,-68% Sales Value is not within the threshold range., -63100.74,,");
		values.add(
				"2634,09-29-2022,success,,,,,,-68% Sales Value is not within the threshold range.,-63100.74,,");
		values.add(
				"263,09-29-2022,failure,,,,,,-68% Sales Value is not within the threshold range.,Calculated EBITDAR Value -63100.74 is negative.  ,,");

		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet(Sheetname);

		XSSFColor orange = new XSSFColor(new java.awt.Color(237, 125, 49));
		Boolean header1bold1 = true;
		XSSFCellStyle header1style1 = getstyle(workbook, orange, header1bold1, defaultfontStyle);

		Row hRow1 = sheet.createRow(0);
		hRow1.setHeight((short) 1500);

		Cell cell1 = hRow1.createCell(0);
		cell1.setCellValue(headers1.get(0).toString());
		cell1.setCellStyle(header1style1);

		XSSFColor blue = new XSSFColor(new java.awt.Color(68, 114, 196));
		Boolean header1bold2 = true;
		XSSFCellStyle header1style2 = getstyle(workbook, blue, header1bold2, defaultfontStyle);

		Cell cell2 = hRow1.createCell(8);
		cell2.setCellValue(headers1.get(1).toString());
		cell2.setCellStyle(header1style2);

		// XSSFColor white = new XSSFColor(new java.awt.Color(255,255,255));
		// XSSFCellStyle header2style1 = getstyle(workbook,blue);

		Row hRow2 = sheet.createRow(rowcount);
		for (int i = 0; i < headers2.size(); i++) {
			Cell cell = hRow2.createCell(i);
			cell.setCellValue(headers2.get(i).toString());
			if (i < 8) {
				cell.setCellStyle(header1style1);
			} else {
				cell.setCellStyle(header1style2);
			}

		}

		

		int valuerowcount = rowcount;

		for (int i = 0; i < values.size(); i++) {
			valuerowcount = valuerowcount + 1;
			Row row = sheet.createRow(valuerowcount);
			String[] datas = values.get(i).split(",");
			System.out.println("data length == "+ datas.length);
			for (int j = 0; j < datas.length; j++) {
				System.out.println(j);
				Cell cell = row.createCell(j);
				XSSFCellStyle valuecellStyle = workbook.createCellStyle();
				setBorder(valuecellStyle);
				valuecellStyle.setAlignment(HorizontalAlignment.CENTER);
				boolean bold = false;
				setFont(workbook , valuecellStyle,defaultfontStyle,bold,defaultfontColor);
				
				if (valuerowcount % 2 == 0) {
					XSSFColor pinkColor = new XSSFColor(new java.awt.Color(252, 228, 214));
					setForegroundColor(valuecellStyle,pinkColor);
				}
				else {
					XSSFColor whiteColor = new XSSFColor(new java.awt.Color(255, 255, 255));
					setForegroundColor(valuecellStyle,whiteColor);
				}
				if (datas[j].equalsIgnoreCase("failure")) {
					
					XSSFCellStyle failurecellStyle = workbook.createCellStyle();
					XSSFColor failurecellColor = new XSSFColor(new java.awt.Color(255, 0, 0));//red color
					setForegroundColor(failurecellStyle,failurecellColor);
					setBorder(failurecellStyle);
					setFont(workbook, failurecellStyle, defaultfontStyle, bold,defaultfontColor);
					cell.setCellValue(datas[j]);
					cell.setCellStyle(failurecellStyle);
				} else {
					
//				if(datas[j].contains("-")) 
//				{
////						_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
//					Font font = workbook.createFont();
//					font.setColor(IndexedColors.RED.getIndex());
//					font.setBold(bold);
//					valuecellStyle.setFont(font);
//					applyFormat(workbook,valuecellStyle);
//					
//					}
					if(datas[j].equals(""))
	                {
						System.out.println("@@@@@@@@@@@");
	                    cell.setCellValue("");
	                    cell.setCellStyle(valuecellStyle);
	                }
					else {
					cell.setCellValue(datas[j]);
					cell.setCellStyle(valuecellStyle);
					}
				}

			}
		}
		if (sheet.getPhysicalNumberOfRows() > 0) {
			for (int i = 0; i <= 1; i++) {
				Row row = sheet.getRow(1);
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					int columnIndex = cell.getColumnIndex();
					sheet.autoSizeColumn(columnIndex);
				}
			}
		}

		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 8, 11));
		sheet.setAutoFilter(new CellRangeAddress(1, 1, 0, headers2.size() - 1));
		sheet.createFreezePane(3, 2);

		FileOutputStream out = new FileOutputStream(path);
		workbook.write(out);
		out.close();
		workbook.close();
	
}
	public static XSSFCellStyle getstyle(XSSFWorkbook workbook, XSSFColor foregroundcolour, boolean bold,
			String fontStyle) {

		XSSFCellStyle style = workbook.createCellStyle();
		style.setWrapText(true);
		style.setAlignment(HorizontalAlignment.CENTER);
		setForegroundColor(style,foregroundcolour);
		setBorder(style);

		short fontColor = IndexedColors.WHITE.getIndex() ;
		setFont(workbook , style,fontStyle,bold,fontColor);
		return style;
	}

	private static void setForegroundColor(XSSFCellStyle style,XSSFColor foregroundcolour) {
		style.setFillForegroundColor(foregroundcolour);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
	}

	public static void setBorder(CellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
	}
	
	public static void setFont(XSSFWorkbook workbook,CellStyle style,String fontStyle,boolean bold,short fontColor) {
	    Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName(fontStyle);
		font.setColor(fontColor);
		font.setBold(bold);
		style.setFont(font);
	}
	public static void applyFormat(XSSFWorkbook workbook, CellStyle style)
	{
		//System.out.println("@@@@@@@@@@@@@@@");
		DataFormat format = workbook.createDataFormat();
		String styleFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
	    style.setDataFormat(format.getFormat(styleFormat));
	}
}
