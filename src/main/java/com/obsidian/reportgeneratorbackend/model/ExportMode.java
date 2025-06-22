package com.obsidian.reportgeneratorbackend.model;

import com.fasterxml.jackson.annotation.JsonProperty;

/*
 * 描述: 定义支持的导出模式的枚举。
 *       使用枚举可以防止传入无效的模式字符串，增强了代码的类型安全。
 */
public enum ExportMode {

    /*
     * @JsonProperty 注解用于将前端传来的 "single-sheet" JSON字符串
     * 自动映射到 Java 的 SINGLE_SHEET 枚举常量上。
     */
    @JsonProperty("single-sheet")
    SINGLE_SHEET, // 合并到单表 (测试用例模式)

    @JsonProperty("multi-sheet")
    MULTI_SHEET, // 合并为多工作簿 (报告模式)

    @JsonProperty("zip-files")
    ZIP_FILES; // 导出为多个独立文件 (报告模式)
}