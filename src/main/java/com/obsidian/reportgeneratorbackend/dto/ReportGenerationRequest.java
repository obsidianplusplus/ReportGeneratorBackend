package com.obsidian.reportgeneratorbackend.dto;

import com.obsidian.reportgeneratorbackend.model.ExportMode;
import lombok.Data;
import java.util.List;
import java.util.Map;

@Data
public class ReportGenerationRequest {

    private ExportMode exportMode;

    private Map<String, SingleCellMapping> mappingRules;

    private List<LogRecord> logData;

    private List<CustomSourceItem> customSourceItems;
}