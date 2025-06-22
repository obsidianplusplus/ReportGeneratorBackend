package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;

/*
 * 描述: 代表一个具体的测试项及其结果值。
 */
@Data
public class DetailedItem {
    /*
     * 测试项目的名称 (例如 "电池电压")。
     * 它将与映射规则的key进行匹配。
     */
    private String itemName;

    /*
     * 该测试项的实际测量值，以字符串形式表示。
     */
    private String actualValue;
}