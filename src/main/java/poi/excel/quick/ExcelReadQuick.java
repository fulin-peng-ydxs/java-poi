package poi.excel.quick;


import org.apache.poi.ss.usermodel.*;
import java.io.IOException;
import java.util.Objects;

/**
 * excel读取案例
 *
 * @author pengshuaifeng
 * 2024/1/3
 */
public class ExcelReadQuick {

    public static void main(String[] args) {
        try {
            //获取xls工作簿对象
            Workbook wb = WorkbookFactory.create(
                    Objects.requireNonNull(ExcelReadQuick.class.getResourceAsStream("/excel/write_quick.xls")));
            //获取第一个sheet
            Sheet sheet = wb.getSheetAt(0);
            //获取第一行
            Row rows = sheet.getRow(0);
            //获取表开始行索引，注意:之前有内容并且后来被设置为空的行可能仍然被Excel和Apache POI计算为行，
            // 因此此方法的结果将包括这样的行，因此返回值可能低于预期
            int firstRowNum = sheet.getFirstRowNum();
            //获取表结束索引，注意:之前有内容并且后来被设置为空的行可能仍然被Excel和Apache POI计算为行，
            //因此此方法的结果将包括这样的行，因此返回值可能低于预期
            int lastRowNum = sheet.getLastRowNum();
            int rowCount=1;
            //遍历每一行
            for (Row row : sheet) {
                //获取第一个单元格
                Cell rowCell = row.getCell(0);
                //获取该行中包含的第一个单元格的索引。注意:之前有内容的单元格，后来被设置为空的单元格可能仍然被Excel和Apache
                // POI计算为单元格，所以这个方法的结果将包括这样的行，因此返回的值可能比预期的要低!
                short firstCellNum = row.getFirstCellNum();
                //获取此行PLUS ONE中包含的最后一个单元格的索引.注意:之前有内容的单元格，后来被设置为空的单元格可能仍然被Excel和Apache
                // POI计算为单元格，所以这个方法的结果将包括这样的行，因此返回的值可能比预期的要高!
                short lastCellNum = row.getLastCellNum();
                System.out.println("遍历第"+rowCount+"行");
                //遍历每一个单元格
                for (Cell cell : row) {
                    CellType cellType = cell.getCellType();
                    System.out.print(( cellType!=CellType.NUMERIC ? cell.getStringCellValue():cell.getNumericCellValue())+"、");
                }
                System.out.println("");
                rowCount++;
            }
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
