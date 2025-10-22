package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class SeriesDefinitionExcel {
    private String name;
    private List<String> dataAddresses;
    private String sheetName;
    private boolean isRange;
    private List<String> comparisonDataAddresses;
}