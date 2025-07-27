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
    private String combinedSheetName;
    private List<SeriesDefinitionExcel> series;

    // [修改后] 新增此字段以接收前端的指令
    private boolean showMinMax;
}