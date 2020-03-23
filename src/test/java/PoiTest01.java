

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * 通过POI完成excel文件的创建
 */
public class PoiTest01 {

    public static void main(String[] args) throws Exception {
        //1.创建工作簿
        /**
         * HSSFWorkbook :    处理2003版本的excel
         * XSSFWorkbook ：   处理2007版本的excel
         * SXSSFWorkbook ：  在2007版本中处理百万数据excel生成
         */
        Workbook wb = new XSSFWorkbook();
        //2.创建页
        Sheet sheet = wb.createSheet("n1");
        //3.创建行
        Row row = sheet.createRow(3);//参数，行索引
        //4.创建单元格
        Cell cell = row.createCell(2);//参数，单元格索引
        //5.在单元格中赋值内容
        cell.setCellValue("黑马125");
        //6.输出Excel文件到硬盘
        FileOutputStream ots = new FileOutputStream("D:\\demo.xlsx");
        wb.write(ots);
        //7.释放资源
        wb.close();
    }
}