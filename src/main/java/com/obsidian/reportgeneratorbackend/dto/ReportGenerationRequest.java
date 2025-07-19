// src/main/java/com/obsidian/reportgeneratorbackend/dto/ReportGenerationRequest.java
package com.obsidian.reportgeneratorbackend.dto;

import com.obsidian.reportgeneratorbackend.model.ExportMode;
import lombok.Data;
import java.util.List;
import java.util.Map;

/*
 * 描述: 封装从前端发送过来的完整报告生成请求。
 *       这是后端API接收的主要数据结构。
 */
@Data
public class ReportGenerationRequest {

    /*
     * 导出模式，例如 'single-sheet' 或 'zip-files'。
     * 使用枚举类型确保了值的有效性。
     */
    private ExportMode exportMode;

    /*
     * 【V9.0 更新】映射规则的数据结构已改变。
     * - Key: 目标单元格地址 (例如: "4_2")
     * - Value: 包含一个源规则列表的映射对象
     */
    private Map<String, SingleCellMapping> mappingRules;

    /*
     * 从前端选中的、需要填充到报告中的日志数据记录列表。
     */
    private List<LogRecord> logData;
}