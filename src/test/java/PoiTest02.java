
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

/**
 * 通过POI完成excel文件的创建
 */
public class PoiTest02 {

    public static void main(String[] args) throws Exception {
        //1.通过已有的excel文件创建工作簿
        Workbook wb = new XSSFWorkbook("G:\\java\\ssm\\ssm项目\\Saas\\day11\\day11-资料\\demo.xlsx");
        //2.获取第一页
        Sheet sheet = wb.getSheetAt(0);//页索引
        //3.循环每一 v行
        //sheet.getLastRowNum() : 获取最后一个有数据的行索引
        for (int i=0;i < sheet.getLastRowNum() + 1; i ++) {
            //获取每一行
            Row row = sheet.getRow(i);
            for (int j=2; j< row.getLastCellNum() ;j ++) { // row.getLastCellNum() ： 获取最受一个有数据的单元格列号
                //4.获取每一行中的每一个单元格
                Cell cell = row.getCell(j);
                //5.获取每个单元格的内容
                Object obj = getCellValue(cell);
                System.out.print(obj + "----");
            }
            System.out.println("");
        }
    }

    //解析每个单元格的数据 : 当前单元格中的数据
    public static Object getCellValue(Cell cell) {
        Object obj = null;
        CellType cellType = cell.getCellType(); //获取单元格数据类型
        switch (cellType) {
            case STRING: { //字符串单元
                obj = cell.getStringCellValue();
                break;
            }
            //excel默认将日志也理解为数字
            case NUMERIC:{ //数字单元格
                if(DateUtil.isCellDateFormatted(cell)) { //日期
                    obj = cell.getDateCellValue();
                }else {
                    obj = cell.getNumericCellValue();
                }
                break;
            }
            case BOOLEAN:{ //boolean
                obj = cell.getBooleanCellValue();
                break;
            }
            default:{
                break;
            }
        }

        return obj;
    }
}