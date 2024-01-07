package poi.utils;


import org.junit.Test;
import poi.excel.entity.User;
import poi.excel.utils.ExcelUtils;
import java.io.FileOutputStream;
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
            Map<String, Collection<User>> users = ExcelUtils.read(ExcelUtilsTest.class.getResourceAsStream("/write_quick.xls"),
                    User.class, null);
            System.out.println(users);
            System.out.println("-------------------------");
            Map<Integer, String> headers = new HashMap<>();
            headers.put(0,"name");
            headers.put(1,"address");
            headers.put(2,"age");
            System.out.println(users);
            users = ExcelUtils.read(ExcelUtilsTest.class.getResourceAsStream("/write_quick.xls"),
                    User.class, headers);
            System.out.println(users);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void  testWriteCustom(){
        try {
            List<User> users = Arrays.asList(new User("test1", "test1-1", 1),
                    new User("test2", "test2-1", 12), new User("test3", "test3-1", 99));
            Map<String, String> headers = new LinkedHashMap<>();
            headers.put("姓名","name");
            headers.put("年龄","age");
            headers.put("地址","address");
            ExcelUtils.write(new FileOutputStream("/Users/pengshuaifeng/write.xlsx"),Collections.singletonMap("sheet-test",users),
                    Collections.singletonMap("sheet-test",headers), ExcelUtils.ExcelType.XLSX,null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void  testWriteModel(){
        try {
            List<User> users = Arrays.asList(new User("test1", "test1-1", 1),
                    new User("test2", "test2-1", 12), new User("test3", "test3-1", 99));
            Map<String, String> headers = new LinkedHashMap<>();
            headers.put("姓名","name");
            headers.put("地址","address");
            headers.put("年龄","age");
            ExcelUtils.write(new FileOutputStream("/Users/pengshuaifeng/write.xls"),
                    ExcelUtils.class.getResourceAsStream("/excel/write_model.xls"),Collections.singletonMap("用户信息列表",users),Collections.singletonMap("用户信息列表",headers),
                    -1,1,0, ExcelUtils.ExcelType.XLS,null);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
