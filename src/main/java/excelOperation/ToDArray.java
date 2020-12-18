package excelOperation;

import com.sun.tools.javac.util.ArrayUtils;
import org.apache.poi.hpsf.Array;
import org.apache.poi.util.ArrayUtil;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.IOException;
import java.util.*;

import static com.sun.tools.javac.util.ArrayUtils.*;

public class ToDArray {
    Object[][] loginData = null;
    HashMap<String,String> data = null;
    public static void main(String[] args) throws IOException {
        System.out.println(Arrays.deepToString(ExcelMap.getTestData("DataDriven.xlsx", "Sheet1")));

    }
    @Test(dataProvider = "getHomeData")
    public void getValue() throws IOException {
        data = new HashMap<String, String>((Integer) ExcelMap.getTestData("DataDriven.xlsx", "Sheet1")[0][0]);
        System.out.println(data);
        loginData = ExcelMap.getTestData("DataDriven.xlsx", "Sheet1");

        System.out.println(Arrays.deepToString(loginData));
//        loginData = ExcelMap.getTestData("DataDriven.xlsx", "Sheet1");
//        System.out.println(loginData[0][0]);
    }
    @Test(dataProvider = "getHomeData")
    public void setData(Map<Object, Object> map) throws Exception {
        System.out.println(map.get("ROWNUM"));
        //ExcelMap.updateTestDataInTestSheet("DataDriven.xlsx","USERNAME","SAMKIT","1","Sheet1");
    }
    @Test
    public void map () throws IOException {

    }



    @DataProvider
    public Object[][] getHomeData() throws Exception {
        return ExcelMap.getTestData("DataDriven.xlsx", "Sheet1");
    }
}
