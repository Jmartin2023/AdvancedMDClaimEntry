package objects;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperations {
    private String path;
    private final XSSFWorkbook wb;
    private XSSFSheet sheet;

    public ExcelOperations(String path) throws IOException {
        this.path = path;

        FileInputStream fis = new FileInputStream(new File(this.path));

        //Create Workbook instance holding reference to .xlsx file
        this.wb = new XSSFWorkbook(fis);

        //Get first/desired sheet from the workbook
        this.sheet = this.wb.getSheetAt(0);
        
        fis.close();
    }

    public ExcelOperations(String path, String sheetName) throws IOException {
        this.path = path;

        FileInputStream file = new FileInputStream(new File(this.path));

        //Create Workbook instance holding reference to .xlsx file
        this.wb = new XSSFWorkbook(file);

        //Get first/desired sheet from the workbook
        this.sheet = this.wb.getSheet(sheetName);
    }

    public void setActiveSheet(String sheetName) throws ExcelExceptions {
        if(this.wb.getSheetIndex(sheetName)==-1) throw new ExcelExceptions("No sheet found");
        this.sheet = this.wb.getSheet(sheetName);
        this.wb.setActiveSheet(this.wb.getSheetIndex(sheetName));
    }

    public int getRowCount(){
        return this.sheet.getLastRowNum();    // add 1 as rows start from 0 index
    }

    public int getColCount(){
        return this.sheet.getRow(0).getLastCellNum();
    }

    public int getColCount(int rowNum){
        return this.sheet.getRow(rowNum).getLastCellNum();
    }

    public String getCellData(int rowNum, int colNum) throws ExcelExceptions {
        if(rowNum<0) throw new ExcelExceptions("Invalid Row Number");
        XSSFRow row = sheet.getRow(rowNum);
        if(row==null) throw new ExcelExceptions("Row does not exist");

        XSSFCell cell = row.getCell(colNum);
        if(cell==null) return "";

        if(cell.getCellType()==CellType.BLANK || cell.getCellType()==CellType._NONE) return "";
        else if(cell.getCellType()==CellType.STRING) return cell.getStringCellValue();
        else if(cell.getCellType()==CellType.NUMERIC) {
            if(DateUtil.isCellDateFormatted(cell)){
                double dateDouble = cell.getNumericCellValue();
                Date date = DateUtil.getJavaDate(dateDouble);
                return this.formatDate(date);
            } else {
                double numVal  = cell.getNumericCellValue();
                if(String.valueOf(numVal).contains("E")) {
                	BigDecimal result = new BigDecimal(String.valueOf(numVal));
                	return result.toPlainString();
                }
                else if ((numVal == Math.floor(numVal)) && !Double.isInfinite(numVal)) {
                    int num = (int) numVal;
                    return String.valueOf(num);
                } else {
                    return String.valueOf(numVal);
                }
            }
            
        } 
        else if(cell.getCellType()==CellType.ERROR) {
        	return cell.getErrorCellString().toString();
        }
        else {
            System.out.println("Cell Type: " + cell.getCellType().toString());
            throw new ExcelExceptions("Invalid Cell Type: " + cell.getCellType().toString());
        }
    }

    public void setCellText(int rowNum, int colNum, String text) throws ExcelExceptions, IOException {

        if(rowNum<0) throw new ExcelExceptions("Invalid Row Number");

        XSSFRow row = sheet.getRow(rowNum);
        if(row==null) sheet.createRow(rowNum);

        XSSFCell cell = row.getCell(colNum);
        if(cell==null) {
            row.createCell(colNum, CellType.STRING);
            cell = row.getCell(colNum);
        }

        cell.setCellValue(text);

        FileOutputStream fos = new FileOutputStream(this.path);
        try {
            this.wb.write(fos);
        }
        finally {
            fos.close();
        }
    }

    private String formatDate(Date date){
        SimpleDateFormat DateFor = new SimpleDateFormat("MM/dd/yyyy");
        String stringDate= DateFor.format(date);
        return stringDate;
    }
    
    public void closeWorkbook() throws IOException {
    	this.wb.close();
    }

    public static class ExcelExceptions extends Exception {
        public ExcelExceptions(String message) {
            super(message);
        }
    }
}
