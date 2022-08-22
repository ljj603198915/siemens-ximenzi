package com.ljj.domain;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import lombok.Data;

@Data
public class Lowshop {
    @TableId(value = "id", type = IdType.AUTO)
    private Long id;

    private String shopName;
    private String province;
    private String city;
    private String district;
    private String adress;
    private String telephone;
    private String longitude;
    private String latitude;
    private Integer shopType;
    private Integer onSale;
    private Integer resTypeId;
    private Integer displayOrder;
}
