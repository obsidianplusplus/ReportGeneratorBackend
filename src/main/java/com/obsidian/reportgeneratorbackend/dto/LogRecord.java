package com.obsidian.reportgeneratorbackend.dto;

import lombok.Data;
import java.util.List;

/*
 * 描述: 代表一条完整的日志记录，通常对应一个SN或一个文件。
 */
@Data
public class LogRecord {
    /*
     * 产品的序列号 (Serial Number)。
     */
    private String sn;

    /*
     * 该记录包含的所有详细测试项的列表。
     */
    private List<DetailedItem> detailedItems;
}