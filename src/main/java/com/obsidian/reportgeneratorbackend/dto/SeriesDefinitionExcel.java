package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class SeriesDefinitionExcel {
    private String name;
    private List<String> dataAddresses;
    private String sheetName; // 新增字段，用于指定数据源所在的工作表名称
}