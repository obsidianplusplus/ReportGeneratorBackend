package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;

@Data
public class CustomSourceItem {
    private String name;
    private String type; // "static" or "random"
    private String value;
    private Double min;
    private Double max;
    private Integer decimals;
}