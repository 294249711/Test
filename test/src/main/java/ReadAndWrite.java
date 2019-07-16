/**
 * Created by 19965 on 2019/7/15.
 */

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * describe:读取Excel数据变成数组形式输出
 *
 * @author lss
 * @date 2019/07/15
 */
public class ReadAndWrite {
    public static List<String> poiRead(List<String> data) throws Exception {
        FileInputStream xlsxStream = null;
        // Excel工作簿 输入流
        xlsxStream = new FileInputStream(new File("D:\\搜狗高速下载\\城区工地分布调研.xlsx"));
        // 构造工作簿对象
        XSSFWorkbook wb = new XSSFWorkbook(xlsxStream);
        // 获取工作表,这里获取的是第一个sheet，
        XSSFSheet sheet = wb.getSheetAt(0);
        String[][] arr = new String[sheet.getLastRowNum() - 1][3];
        System.out.println(sheet.getLastRowNum());
        //循环读出每条记录，第0,1行为标题行，故从下标为2的行开始取数值
        for (int i = 2; i <= sheet.getLastRowNum(); i++) {
            // 获取行,行号作为参数传递给getRow方法
            XSSFRow row = sheet.getRow(i);
            // 获取单元格,row已经确定了行号,列号作为参数传递给getCell，就可以获得相应的单元格了
            XSSFCell codeCell = row.getCell(0);
            // 获取单元格的值
            String code1 = row.getCell(1).getStringCellValue();
            arr[i - 2][0] = code1;
            String code2 = row.getCell(2).getStringCellValue();
            String[] parts = code2.split(",");
            arr[i - 2][1] = parts[0];
            arr[i - 2][2] = parts[1];
            data.add("[\"" + arr[i - 2][0] + "\"," + arr[i - 2][1] + "," + arr[i - 2][2] + "]");
        }
        return data;
    }

    public static void poiWrite(List<String> data) throws Exception {
        File f = new File("D:\\result.txt");
        FileOutputStream fos1 = new FileOutputStream(f);
        OutputStreamWriter dos1 = new OutputStreamWriter(fos1);
        dos1.write("[");
        for (int i = 0; i < data.size(); i++) {
            dos1.write(data.get(i));
            if (i != data.size() - 1) {
                dos1.write(",");
            }
        }
        dos1.write("]");
        dos1.close();
    }

    public static void main(String[] args) throws Exception {
        List<String> data = new ArrayList<>();
        data = poiRead(data);
        poiWrite(data);
    }
}
