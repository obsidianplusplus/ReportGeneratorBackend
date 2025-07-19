// src/main/java/com/obsidian/reportgeneratorbackend/dto/SourceRule.java
package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;

@Data
public class SourceRule {
    private String sourceKey;
    private String unit;
    private Integer decimals;
}