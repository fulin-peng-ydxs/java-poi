package poi.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * user
 *
 * @author pengshuaifeng
 * 2024/1/4
 */
@Data
@AllArgsConstructor
public class User {
    private String name;
    private String address;
    private int age;

}
