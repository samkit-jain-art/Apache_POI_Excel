package excelOperation;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadingExcelFile {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "DataFile/DataDriven.xlsx";

        FileInputStream fileInputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
      XSSFSheet xssfSheet = xssfWorkbook.getSheet("Sheet1");

      int rows = xssfSheet.getLastRowNum();
      int cols = xssfSheet.getRow(1).getLastCellNum();
      for (int r = 0;r<=rows;r++){
          XSSFRow row = xssfSheet.getRow(r);
          for (int c = 0;c<cols;c++){
           XSSFCell cell= row.getCell(c);
         switch (cell.getCellType())
              {
                  case STRING:
                      System.out.println(cell.getStringCellValue());
                      break;
                  case NUMERIC:
                      System.out.println(cell.getNumericCellValue());
                      break;
                  case BOOLEAN:
                      System.out.println(cell.getBooleanCellValue());
              }
          }
          System.out.println();
      }

    }
}
