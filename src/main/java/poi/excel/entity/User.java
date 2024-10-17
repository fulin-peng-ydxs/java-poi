package poi.excel.entity;

import io.swagger.annotations.ApiModelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

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
    @ApiModelProperty("姓名")
    private String name;
    @ApiModelProperty("地址")
    private String address;
    @ApiModelProperty("年龄")
    private int age;
    @ApiModelProperty("登记时间")
    private Date registerTime;
}
