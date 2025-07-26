package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class SeriesDefinitionExcel {
    private String name;
    private List<String> dataAddresses;
    private String sheetName;
    private boolean isRange; // 新增字段，标记数据源是否为范围
}