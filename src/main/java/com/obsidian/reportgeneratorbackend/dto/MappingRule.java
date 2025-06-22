package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;

/*
 * 描述: 定义了单个数据项如何映射到Excel单元格的规则。
 */
@Data
public class MappingRule {
    /*
     * 目标单元格地址，格式为 "行_列" (例如 "4_2" 表示第5行C列)。
     * 这是从 Luckysheet 直接获取的格式。
     */
    private String address;

    /*
     * 附加的单位 (例如 "V", "A")。如果为null则不添加。
     */
    private String unit;

    /*
     * 要保留的小数位数。如果为null则不进行格式化。
     */
    private Integer decimals;
}