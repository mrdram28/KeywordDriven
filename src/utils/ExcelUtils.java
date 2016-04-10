package utils;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	private String filepath;
	
	private XSSFWorkbook workbook;
	
	// private static Logger logger=null;
	private InputStream resourceAsStream;

	public ExcelUtils(String filePath) {
	    // logger=LoggerFactory.getLogger("ReadXLSX");
	    this.filepath = filePath;
	    resourceAsStream = ClassLoader.getSystemResourceAsStream(filepath);
	}

	public ExcelUtils(InputStream fileStream) {
	    // logger=LoggerFactory.getLogger("ReadXLSX");
	    this.resourceAsStream = fileStream;
	}

	private void loadFile() throws FileNotFoundException
	         {

	    if (resourceAsStream == null)
	        throw new FileNotFoundException("Unable to locate give file..");
	    else {
	        try {
	            workbook = new XSSFWorkbook(resourceAsStream);

	        } catch (IOException ex) {

	        }

	    }
	}// end loadxlsFile

	public String[] getSheetsName() {
	    int totalsheet = 0;
	    int i = 0;
	    String[] sheetName = null;

	    try {
	        loadFile();
	        totalsheet = workbook.getNumberOfSheets();
	        sheetName = new String[totalsheet];
	        while (i < totalsheet) {
	            sheetName[i] = workbook.getSheetName(i);
	            i++;
	        }

	    } catch (FileNotFoundException ex) {
	        // logger.error(ex);
	    } catch (Exception ex) {
	        // logger.error(ex);
	    }

	    return sheetName;
	}

	public int[] getSheetsIndex() {
	    int totalsheet = 0;
	    int i = 0;
	    int[] sheetIndex = null;
	    String[] sheetname = getSheetsName();
	    try {
	        loadFile();
	        totalsheet = workbook.getNumberOfSheets();
	        sheetIndex = new int[totalsheet];
	        while (i < totalsheet) {
	            sheetIndex[i] = workbook.getSheetIndex(sheetname[i]);
	            i++;
	        }

	    } catch (FileNotFoundException ex) {
	        // logger.error(ex);
	    } catch (Exception ex) {
	        // logger.error(ex);
	    }

	    return sheetIndex;
	}

	private boolean validateIndex(int index) {
	    if (index < getSheetsIndex().length && index >= 0)
	        return true;
	    else
	        return false;
	}

	public int getNumberOfSheet() {
	    int totalsheet = 0;
	    try {
	        loadFile();
	        totalsheet = workbook.getNumberOfSheets();

	    } catch (FileNotFoundException ex) {
	        // logger.error(ex.getMessage());
	    } catch (Exception ex) {
	        // logger.error(ex.getMessage());
	    }

	    return totalsheet;
	}

	public int getNumberOfColumns(int SheetIndex) {
	    int NO_OF_Column = 0;
	    @SuppressWarnings("unused")
	    XSSFCell cell = null;
	    XSSFSheet sheet = null;
	    try {
	        loadFile(); // load give Excel
	        if (validateIndex(SheetIndex)) {
	            sheet = workbook.getSheetAt(SheetIndex);
	            Iterator<Row> rowIter = sheet.rowIterator();
	            XSSFRow firstRow = (XSSFRow) rowIter.next();
	            Iterator<Cell> cellIter = firstRow.cellIterator();
	            while (cellIter.hasNext()) {
	                cell = (XSSFCell) cellIter.next();
	                NO_OF_Column++;
	            }
	        } else
	        	
	            throw new Exception("Invalid sheet index.");
	        
	    } catch (Exception ex) {
	        // logger.error(ex.getMessage());

	    }

	    return NO_OF_Column;
	}

	public int getNumberOfRows(int SheetIndex) {
	    int NO_OF_ROW = 0;
	    XSSFSheet sheet = null;

	    try {
	        loadFile(); // load give Excel
	        if (validateIndex(SheetIndex)) {
	            sheet = workbook.getSheetAt(SheetIndex);
	            NO_OF_ROW = sheet.getLastRowNum();
	        } else
	        	
	            throw new Exception("Invalid sheet index.");
	        
	    } catch (Exception ex) {
	        // logger.error(ex);
	    }

	    return NO_OF_ROW;
	}

	public String[] getSheetHeader(int SheetIndex) {
	    int noOfColumns = 0;
	    XSSFCell cell = null;
	    int i = 0;
	    String columns[] = null;
	    XSSFSheet sheet = null;

	    try {
	        loadFile(); // load give Excel
	        if (validateIndex(SheetIndex)) {
	            sheet = workbook.getSheetAt(SheetIndex);
	            noOfColumns = getNumberOfColumns(SheetIndex);
	            columns = new String[noOfColumns];
	            Iterator<Row> rowIter = sheet.rowIterator();
	            XSSFRow Row = (XSSFRow) rowIter.next();
	            Iterator<Cell> cellIter = Row.cellIterator();

	            while (cellIter.hasNext()) {
	                cell = (XSSFCell) cellIter.next();
	                columns[i] = cell.getStringCellValue();
	                i++;
	            }
	        } else
	            throw new Exception("Invalid sheet index.");
	    }

	    catch (Exception ex) {
	        // logger.error(ex);
	    }

	    return columns;
	}// end of method

	public String[][] getSheetData(int SheetIndex) {
	    int noOfColumns = 0;
	    XSSFRow row = null;
	    XSSFCell cell = null;
	    int i = 0;
	    int noOfRows = 0;
	    int j = 0;
	    String[][] data = null;
	    XSSFSheet sheet = null;

	    try {
	        loadFile(); // load give Excel
	        if (validateIndex(SheetIndex)) {
	            sheet = workbook.getSheetAt(SheetIndex);
	            noOfColumns = getNumberOfColumns(SheetIndex);
	            noOfRows = getNumberOfRows(SheetIndex) + 1;
	            data = new String[noOfRows][noOfColumns];
	            Iterator<Row> rowIter = sheet.rowIterator();
	            while (rowIter.hasNext()) {
	                row = (XSSFRow) rowIter.next();
	                Iterator<Cell> cellIter = row.cellIterator();
	                j = 0;
	                while (cellIter.hasNext()) {
	                    cell = (XSSFCell) cellIter.next();
	                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                        data[i][j] = cell.getStringCellValue();
	                    } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
	                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
	                            String formatCellValue = new DataFormatter()
	                                    .formatCellValue(cell);
	                            data[i][j] = formatCellValue;
	                        } else {
	                            data[i][j] = Double.toString(cell
	                                    .getNumericCellValue());
	                        }

	                    } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
	                        data[i][j] = Boolean.toString(cell
	                                .getBooleanCellValue());
	                    }

	                    else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
	                        data[i][j] = cell.getCellFormula().toString();
	                    }

	                    j++;
	                }

	                i++;
	            } // outer while

	        } else
	            throw new Exception("Invalid sheet index.");

	    } catch (Exception ex) {
	        // logger.error(ex);
	    }
	    return data;
	}

	public String[][] getSheetData(int SheetIndex, int noOfRows) {
	    int noOfColumns = 0;
	    XSSFRow row = null;
	    XSSFCell cell = null;
	    int i = 0;
	    int j = 0;
	    String[][] data = null;
	    XSSFSheet sheet = null;

	    try {
	        loadFile(); // load give Excel

	        if (validateIndex(SheetIndex)) {
	            sheet = workbook.getSheetAt(SheetIndex);
	            noOfColumns = getNumberOfColumns(SheetIndex);
	            data = new String[noOfRows][noOfColumns];
	            Iterator<Row> rowIter = sheet.rowIterator();
	            while (i < noOfRows) {

	                row = (XSSFRow) rowIter.next();
	                Iterator<Cell> cellIter = row.cellIterator();
	                j = 0;
	                while (cellIter.hasNext()) {
	                    cell = (XSSFCell) cellIter.next();
	                    if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                        data[i][j] = cell.getStringCellValue();
	                    } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
	                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
	                            String formatCellValue = new DataFormatter()
	                                    .formatCellValue(cell);
	                            data[i][j] = formatCellValue;
	                        } else {
	                            data[i][j] = Double.toString(cell
	                                    .getNumericCellValue());
	                        }
	                    }

	                    j++;
	                }

	                i++;
	            } // outer while
	        } else
	            throw new Exception("Invalid sheet index.");
	    } catch (Exception ex) {
	        // logger.error(ex);
	    }

	    return data;
	}

}

