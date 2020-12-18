package excelOperation;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

public class ExcelToHasMap {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "DataFile/StudentDriven.xlsx";
        FileInputStream fileInputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheet("Student Data");
        int rows = xssfSheet.getLastRowNum();
        HashMap<String,String> data = new HashMap<String, String>();
        for (int r = 0;r<=rows;r++){
           String key = xssfSheet.getRow(r).getCell(0).getStringCellValue();
            String value = xssfSheet.getRow(r).getCell(1).getStringCellValue();
        }
    }
}
