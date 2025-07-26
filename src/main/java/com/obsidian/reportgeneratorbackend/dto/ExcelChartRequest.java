package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class ExcelChartRequest {
    private String chartType;
    private String title;
    private String xAxisTitle;
    private String yAxisTitle;
    private String outputMode; // 新增字段, e.g., "combined" or "separate"
    private List<SeriesDefinitionExcel> series;
}