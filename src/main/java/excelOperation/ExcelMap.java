package excelOperation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelMap {
    /*
     * To give excel sheet name with index no. of sheet
     * sheet will store under resource folder//src/test/resources/DataDriven.xlsx
     */

    public static Object[][] getTestData(String excelName, String SheetName) throws IOException {
        File file = new File(System.getProperty("user.dir")+"/src/test/resources/"+excelName);
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet xssfSheet = workbook.getSheet(SheetName);
        workbook.close();
        int lastRowNum = xssfSheet.getLastRowNum();
        int lastCellNum = xssfSheet.getRow(0).getLastCellNum();
        List<Map<String, String>> dataMap = getRowData(lastRowNum, lastCellNum, xssfSheet);
        Object[][] objects = new Object[dataMap.size()][1];
        for (int i = 0; i < dataMap.size(); i++) {
            objects[i][0] = dataMap.get(i);
        }
        return objects;
    }

    /*
     *  return excel test data into list of map
     */
    private static List<Map<String, String>> getRowData(int lastRowNum, int lastCellNum, XSSFSheet xssfSheet) {
        List<Map<String, String>> listOfTestData = new ArrayList<Map<String, String>>();
        for (int i = 0; i < lastRowNum; i++) {
            Map<String, String> dataMap = new LinkedHashMap<String, String>();
            if (xssfSheet.getRow(i + 1).getCell(0).toString().equalsIgnoreCase("true")) {
                for (int j = 0; j < lastCellNum; j++) {
                    dataMap.put(xssfSheet.getRow(0).getCell(j).toString(), xssfSheet.getRow(i + 1).getCell(j).toString());
                }
            } else {
                continue;
            }
            listOfTestData.add(dataMap);
        }
        return listOfTestData;
    }
    /*
     * it will update the test data in excel sheet
     * excelName give sheet name
     * @Params key: keyName(column name)
     * value: What value need to update against that key
     * dataRow: pass row number
     * int sheetNo: pass sheetNo;
     * dataRow: pass row number
     * int sheetNo : pass sheetNo.
     */

    public static void updateTestDataInTestSheet(String excelName, String key, String value, String dataRow, String sheetName) throws Exception {
        File file = new File(System.getProperty("user.dir") + "/src/test/resources/" + excelName);
        FileInputStream fileInputStream = new FileInputStream(file);
        @SuppressWarnings("resource")
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet xssfSheet = xssfWorkbook.getSheet(sheetName);
        int lastRowNum = xssfSheet.getLastRowNum();
        int lastCellNum = xssfSheet.getRow(0).getLastCellNum();
        xssfSheet.autoSizeColumn(lastCellNum);
        for (int i = 0; i < lastRowNum; i++) {
            if (xssfSheet.getRow(i + 1).getCell(1).toString().equalsIgnoreCase(dataRow)) {
                try {
                    for (int j = 0; j < lastCellNum; j++) {
                        if (xssfSheet.getRow(0).getCell(j).toString().equalsIgnoreCase(key)) {
                            Cell cell = xssfSheet.getRow(i + 1).getCell(j);
                            if (cell == null) {
                                cell = xssfSheet.getRow(i + 1).createCell(j);
                            }
                            //  cell.setCellType();
                            cell.setCellValue(value);
                        }
                    }
                } catch (Exception e) {
                    throw new Exception("Problem While setting data @rowNUM= " + dataRow + "and for key " + key);
                }

            } else {
                continue;
            }
        }
    }

}



