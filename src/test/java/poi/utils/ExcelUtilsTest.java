package poi.utils;


import org.junit.Test;
import poi.excel.entity.User;
import poi.excel.utils.ExcelUtils;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

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
            Collection<User> users = ExcelUtils.read(ExcelUtilsTest.class.getResourceAsStream("/write_quick.xls"),
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
}
