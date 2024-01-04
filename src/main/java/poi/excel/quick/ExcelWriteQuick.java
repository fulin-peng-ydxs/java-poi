package poi.excel.quick;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import poi.excel.entity.User;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedList;

/**
 * excel导出案例
 *
 * @author pengshuaifeng
 * 2024/1/2
 */
@Slf4j
public class ExcelWriteQuick {

    public static void main(String[] args){
        try {
            //创建xls工作簿对象
            HSSFWorkbook wb = new HSSFWorkbook();
            //创建一个sheet
            Sheet s = wb.createSheet("用户信息列表");
            LinkedList<User> users = new LinkedList<>(Arrays.asList(new User("用户1", "用户地址-1", 1),
                    new User("用户2", "用户地址-2", 12), new User("用户3", "用户地址-3", 99)));
            //创建第一行
            Row row = s.createRow(0);
            Cell userNameHead = row.createCell(0);
            userNameHead.setCellValue("用户名");
            Cell addressHead = row.createCell(1);
            addressHead.setCellValue("地址");
            Cell ageHead = row.createCell(2);
            ageHead.setCellValue("年龄");
            int rowIndex=1;
            for (User user : users) {
                Row rowData = s.createRow(rowIndex);
                Cell userName = rowData.createCell(0);
                userName.setCellValue(user.getName());
                Cell address = rowData.createCell(1);
                address.setCellValue(user.getAddress());
                Cell age = rowData.createCell(2);
                age.setCellValue(user.getAge());
                rowIndex++;
            }
            //生成excel文件
            FileOutputStream outputStream = new FileOutputStream("/Users/pengshuaifeng/IdeaProjects/java-poi/src/main/resources/excel/write_quick.xls");
            wb.write(outputStream);
            outputStream.close();
        } catch (IOException e) {
            log.error("excel写入异常",e);
        }
    }
}
