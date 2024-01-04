package poi.excel.utils;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.util.*;
/**
 * excel工具类
 *
 * @author pengshuaifeng
 * 2024/1/4
 */
public class ExcelUtils {

    /**
     * 读取excel
     * <p>表头默认为第一行</p>
     * @param in excel输入流
     * @param targetType 读取实体对象
     * @param headers 工作表表头索引映射，为空则会按照默认规则自动生成
     * @param startReadRow 读取表开始行索引
     * @param startReadCol 读取表开始列索引
     * 2024/1/4 22:23
     * @author pengshuaifeng
     */
    public static <T> Collection<T> read(InputStream in,Class<T> targetType,Map<Integer,String> headers,int startReadRow, int startReadCol) throws Exception {
        return read(in, targetType, -1, headers, 0, startReadRow, startReadCol);
    }

    /**
     * 读取excel
     * <p>表头默认为第一行,从第二行第一列开始读取</p>
     * 2024/1/4 22:23
     * @author pengshuaifeng
     */
    public static <T> Collection<T> read(InputStream in,Class<T> targetType,Map<Integer,String> headers) throws Exception {
        return read(in, targetType, -1, headers, 0, 1, 0);
    }

    /**
     * 读取excel
     * <p>表头默认为第一行</p>
     * @param in excel输入流
     * @param targetType 读取实体对象
     * @param sheetAt 读取的工作表索引，为-1，则会读取所有的表
     * @param headers 工作表表头索引映射，为空则会按照默认规则自动生成
     * @param startReadRow 读取表开始行索引
     * @param startReadCol 读取表开始列索引
     * 2024/1/4 22:23
     * @author pengshuaifeng
     */
    public static <T> Collection<T> read(InputStream in,Class<T> targetType,int sheetAt,Map<Integer,String> headers,int startReadRow, int startReadCol) throws Exception {
        return read(in, targetType, sheetAt, headers, 0, startReadRow, startReadCol);
    }

    /**
     * 读取excel
     * <p>表头默认为第一行,从第二行第一列开始读取</p>
     * 2024/1/4 22:23
     * @author pengshuaifeng
     */
    public static <T> Collection<T> read(InputStream in,Class<T> targetType,int sheetAt,Map<Integer,String> headers) throws Exception {
        return read(in, targetType, sheetAt, headers, 0, 1, 0);
    }

    /**
     * 读取excel
     * @param in excel输入流
     * @param targetType 读取实体对象
     * @param sheetAt 读取的工作表索引，为-1，则会读取所有的表
     * @param headers 工作表表头索引映射，为空则会按照默认规则自动生成
     * @param headerAt 工作表表头所在行索引，headers为空时使用
     * @param startReadRow 读取表开始行索引
     * @param startReadCol 读取表开始列索引
     * 2024/1/4 01:16
     * @author pengshuaifeng
     */
    public static <T> Collection<T> read(InputStream in,Class<T> targetType,int sheetAt,Map<Integer,String> headers,int headerAt,int startReadRow, int startReadCol)throws Exception{
        Workbook workbook = generateWorkBook(in);
        return readWorkbook(workbook,targetType,sheetAt,headers,headerAt,startReadRow,startReadCol);
    }


    /**
     * 读取工作簿
     * 2024/1/4 23:41
     * @author pengshuaifeng
     */
    private static  <T> Collection<T> readWorkbook(Workbook workbook,Class<T> targetType,int sheetAt,Map<Integer,String> headers,int headerAt,int startReadRow, int startReadCol){
        try {
            //读取结果集
            Collection<T> results=null;
            if(sheetAt==-1){ //遍历读取所有工作表
                results=new LinkedList<>();
                for (Sheet sheet : workbook) {
                    headers=getHeaders(headers,sheet,headerAt,targetType); //定义表头
                    results.addAll(readSheet(sheet,headers,targetType,startReadRow,startReadCol)); //读取工作表
                }
            }else{   //读取指定工作表
                Sheet sheet = workbook.getSheetAt(sheetAt);
                headers=getHeaders(headers,sheet,headerAt,targetType);//定义表头
                results=readSheet(sheet,headers,targetType,startReadRow,startReadCol);//读取工作表
            }
            return results;
        } catch (Exception e) {
           throw new RuntimeException("读取工作簿异常",e);
        }
    }

    /**
     * 读取工作表
     * 2024/1/4 23:41
     * @author pengshuaifeng
     */
    private static  <T> Collection<T> readSheet(Sheet sheet,Map<Integer,String> headers,Class<T> targetType, int startReadRow, int startReadCol){
        try {
            Collection<T> results = new LinkedList<>();
            for (Row row : sheet) {
                if (row.getRowNum()<startReadRow) //从startReadRow索引行开始读取
                    continue;
                results.add(readRow(row,headers,targetType,startReadCol));
            }
            return results;
        } catch (Exception e) {
            throw new RuntimeException("读取工作表异常",e);
        }
    }

    /**
     * 读取数据行
     * 2024/1/4 23:40
     * @author pengshuaifeng
     */
    private static  <T> T readRow(Row row, Map<Integer,String> headers, Class<T> targetType, int startReadCol){
        try {
            T t = targetType.newInstance();
            for (Cell cell : row) {
                int columnIndex = cell.getColumnIndex();
                if (columnIndex <startReadCol) //从startReadCol索引列开始读取
                    continue;
                Field field = targetType.getDeclaredField(headers.get(columnIndex));
                field.setAccessible(true);
                field.set(t,readCell(cell,field.getType()));  //读取单元格数据
            }
            return t;
        } catch (Exception e) {
            throw new RuntimeException("读取数据行异常",e);
        }
    }


    /**
     * 读取单元格
     * 2024/1/4 23:35
     * @author pengshuaifeng
     * @param cell 单元格对象
     * @param targetType 目标数据类型
     */
    //TODO 数字类型读取待完善
    private static Object readCell(Cell cell, Class<?> targetType){
        switch (cell.getCellType()) {
            case STRING:  //字符类型
                return cell.getRichStringCellValue().getString();
            case NUMERIC:  //数字类型
                if (DateUtil.isCellDateFormatted(cell)) {  //时间类型处理
                    if (targetType.isAssignableFrom(LocalDateTime.class)) {
                        return cell.getLocalDateTimeCellValue();
                    }
                    return cell.getDateCellValue();
                } else {
                    double numericCellValue = cell.getNumericCellValue();
                    if (targetType.isAssignableFrom(int.class) || targetType.isAssignableFrom(Integer.class) ) {
                        //int类型处理
                        return  (int)numericCellValue;
                    }else{
                        //其他类型处理
                        return numericCellValue;
                    }
                }
            case BOOLEAN:  //布尔类型
                return cell.getBooleanCellValue();
            case FORMULA:  //公式类型：转换成字符表达式
                return cell.getCellFormula();
            case BLANK:  //空白类型
                if(targetType.isAssignableFrom(Number.class)||targetType.isAssignableFrom(int.class)){
                    return 0;  //数字类型处理
                }else{  //其他类型处理
                    return null;
                }
            default:
                throw new RuntimeException("单元格<"+cell.getRowIndex()+"/"+cell.getColumnIndex()+">未知的数据类型："+cell.getCellType());
        }
    }


    /**
     * 导出excel
     * 2024/1/4 01:17
     * @author pengshuaifeng
     */


    /**
     * 生成工作簿
     * @param in excel输入流
     * 2024/1/4 20:50
     * @author pengshuaifeng
     */
    private static Workbook generateWorkBook(InputStream in){
        try {
            return WorkbookFactory.create(in);
        } catch (IOException e) {
            throw new RuntimeException("打开工作簿异常",e);
        }
    }

    /**
     * 生成表头
     * @param sheet 工作表
     * @param index 表头所在的行
     * @param source 表头映射源
     * 2024/1/4 21:37
     * @author pengshuaifeng
     */
    private static Map<Integer,String> generateHeaders(Sheet sheet,int index,Class<?> source){
        Map<Integer, String> headers = new HashMap<>();
        Row row = sheet.getRow(index);
        //TODO 默认只支持指定表头行为属性名的，后续可根据属性的注解获取符合表头标识符的进行映射
        for (Cell cell : row) {
            int columnIndex = cell.getColumnIndex();
            String header = cell.getStringCellValue();
            headers.put(columnIndex,header);
        }
        return headers;
    }

    /**
     * 获取表头
     * @param headers 表头对象，为空则自动生成表头
     * 2024/1/4 22:42
     * @author pengshuaifeng
     */
    private static Map<Integer,String> getHeaders(Map<Integer,String> headers,Sheet sheet,int index,Class<?> source){
        if (headers==null || headers.isEmpty()) {
            headers=generateHeaders(sheet,index,source);
        }
        return headers;
    }

    /**
     * excel类型
     */
    public enum ExcelType{
        XLS,
        XLSX
    }
}
