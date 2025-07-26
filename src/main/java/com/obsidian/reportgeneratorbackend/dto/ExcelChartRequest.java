package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class ExcelChartRequest {
    private String chartType;
    private String title;
    private String xAxisTitle;
    private String yAxisTitle;
    private List<SeriesDefinitionExcel> series;
}