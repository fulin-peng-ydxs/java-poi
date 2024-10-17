package poi.utils;


import org.junit.Test;
import poi.excel.entity.User;
import poi.excel.utils.ExcelUtils;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

/**
 * excel工具类测试
 *
 * @author pengshuaifeng
 * 2024/1/4
 */
public class ExcelUtilsTest {

    @Test
    public void testRead(){
        try {
            Collection<User> users =ExcelUtils.getDefaultResultCollection(  ExcelUtils.read(ExcelUtilsTest.class.getResourceAsStream("/write_quick.xls"),
                    User.class, null));
            System.out.println(users);
            System.out.println("-------------------------");
            Map<Object, String> headers = new HashMap<>();
            headers.put("用户名","name");
            headers.put("地址","address");
            headers.put("年龄","age");
            System.out.println(users);
            users = ExcelUtils.getDefaultResultCollection(
                    ExcelUtils.read(ExcelUtilsTest.class.getResourceAsStream("/write_read.xls"), User.class, headers));
            System.out.println(users);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

//    @Test
    public void  testWriteCustom(){
        try {
            Date date = new Date();
            List<User> users = Arrays.asList(new User("test1", "test1-1", 1,date),
                    new User("test2", "test2-1", 12,date),
                    new User("test3", "test3-11111111111111111111111111", 99,date));
            Map<String, String> headers = new LinkedHashMap<>();
            headers.put("姓名","name");
            headers.put("年龄","age");
            headers.put("地址","address");
            ExcelUtils.write(Files.newOutputStream(Paths.get("/Users/pengshuaifeng/write.xlsx")),Collections.singletonMap("sheet-test",users),
                    Collections.singletonMap("sheet-test",headers), ExcelUtils.ExcelType.XLSX,null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void  testWriteSwagger(){
        try {
            Date date = new Date();
            List<User> users = Arrays.asList(new User("test1", "test1-1", 1,date),
                    new User("test2", "test2-1", 12,date),
                    new User("test3", "test3-11111111111111111111111111", 99,date));
            ExcelUtils.ExcelCellStyleModel.DEFAULT_HEADER_STYLE.setColumnWidth(4000);
            ExcelUtils.write(Files.newOutputStream(Paths.get("F:\\write.xlsx")),Collections.singletonMap("sheet-test",users),
                    Collections.singletonMap("sheet-test",ExcelUtils.generateHeaders(User.class,Arrays.asList("name","address","age","registerTime"),null)), ExcelUtils.ExcelType.XLSX,null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

//    @Test
    public void  testWriteModel(){
        try {
            Date date = new Date();
            List<User> users = Arrays.asList(new User("test1", "test1-1", 1,date),
                    new User("test2", "test2-1", 12,date), new User("test3", "test3-1", 99,date));
            Map<String, String> headers = new LinkedHashMap<>();
            headers.put("姓名","name");
            headers.put("地址","address");
            headers.put("年龄","age");
            ExcelUtils.write(Files.newOutputStream(Paths.get("/Users/pengshuaifeng/write_model.xlsx")),
                    ExcelUtils.class.getResourceAsStream("/excel/write_model.xls"),Collections.singletonMap("用户信息列表",users),Collections.singletonMap("用户信息列表",headers),
                    -1,1,0, ExcelUtils.ExcelType.XLS,null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void  testWriteModelReplace(){
        try {
            Map<String, Object> data = new HashMap<>();
            data.put("user1","傻逼1");
            data.put("age1","年龄1");
            data.put("sex1","性别1");
            data.put("user2","傻逼2");
            data.put("age2","年龄2");
            data.put("sex2","性别2");
            ExcelUtils.write(Files.newOutputStream(Paths.get("E:\\java-poi\\src\\test\\demo.xls")),
                    ExcelUtils.class.getResourceAsStream("/excel/write_model_replace.xls"),
                    Collections.singletonMap("用户信息列表",data),null,1000,0,0,
                    ExcelUtils.ExcelType.XLS,null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
