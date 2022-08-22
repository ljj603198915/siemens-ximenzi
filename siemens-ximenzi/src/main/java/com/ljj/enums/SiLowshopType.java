package com.ljj.enums;

public enum SiLowshopType {
    SQD("授权店", 1), XXD("形象店", 2);

    private final String shopTypeStr;
    private final Integer shopType;

    SiLowshopType(String shopTypeStr, Integer shopType) {
        this.shopTypeStr = shopTypeStr;
        this.shopType = shopType;
    }

    public String getShopTypeStr() {
        return shopTypeStr;
    }

    public Integer getShopType() {
        return shopType;
    }
}
