package poi.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * user
 *
 * @author pengshuaifeng
 * 2024/1/4
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class User {
    private String name;
    private String address;
    private int age;

}
