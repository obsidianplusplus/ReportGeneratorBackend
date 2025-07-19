// src/main/java/com/obsidian/reportgeneratorbackend/dto/SingleCellMapping.java
package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

@Data
public class SingleCellMapping {
    private List<SourceRule> sources;
}