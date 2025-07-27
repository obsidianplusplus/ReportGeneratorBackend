package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class ExcelChartRequest {
    private String chartType;
    private String title;
    private String xAxisTitle;
    private String yAxisTitle;
    private String outputMode;

    // [核心修改] 新增此字段以接收前端传来的自定义工作表名称
    private String combinedSheetName;

    private List<SeriesDefinitionExcel> series;
}