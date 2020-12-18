package excelOperation;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class HasMapToExcel {
    public static void main(String[] args) throws IOException {

        wrightTheExcelFile();

    }

    public static void wrightTheExcelFile() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet("EMP INFO");
        Object[][] data = {
                {"EmpId", "EmpName", "Job"},
                {"101", "Samkit", "QA"}
        };
        int row = data.length;
        int col = data[0].length;

        System.out.println(row);
        System.out.println(col);
        for (int i = 0; i < row; i++) {
            XSSFRow rows = xssfSheet.createRow(i);
            for (int c = 0; c < col; c++) {
                XSSFCell cell = rows.createCell(c);
                Object value = data[i][c];

                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);


            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream("DataFile/Student.xlsx");
        xssfWorkbook .write(fileOutputStream);
    }
    public static void createData() throws IOException {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet =  xssfWorkbook.createSheet("Student Data");

        Map<String,String> data = new HashMap<String,String>();
        data.put("101","Jon");
        data.put("102","Jon");
        data.put("103","Jon");
        data.put("104","Jon");
        data.put("105","Jon");

        int rowNo=0;
        for (Map.Entry entry:data.entrySet()){
            XSSFRow xssfRow = xssfSheet.createRow(rowNo++);

            xssfRow.createCell(0).setCellValue((String) entry.getKey());
            xssfRow.createCell(1).setCellValue((String) entry.getValue());
        }
        String excelFilePath = "DataFile/StudentDriven.xlsx";
        FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath);
        xssfWorkbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Excel written successfully");

    }
}
