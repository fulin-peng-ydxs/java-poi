package poi.utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.*;

/**
 * excel测试
 *
 * @author pengshuaifeng
 * 2025/5/15
 */
public class ExcelTest {

    @Test
    public void  writeImageExcel(){
        try {
            // 1. 创建工作簿
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("图片示例");

            // 2. 读取图片为字节数组
            InputStream is = ExcelTest.class.getResourceAsStream("/test.png");
            byte[] bytes = new byte[is.available()];
            is.read(bytes);
            int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            is.close();

            // 3. 创建绘图管理器
            CreationHelper helper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            // 4. 设置图片位置（锚点）：从第2行第2列开始（索引从0开始）
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(0); // 列起点
            anchor.setRow1(0); // 行起点
            anchor.setCol2(1); // 列终点
            anchor.setRow2(1); // 行终点（决定图片大小）

            // 5. 插入图片
            drawing.createPicture(anchor, pictureIdx);


            // 6. 保存到文件
            FileOutputStream fileOut = new FileOutputStream("/Users/pengshuaifeng/IdeaProjects/java-poi/src/test/resources/imageOutput.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("图片插入成功！");


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
