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
     * 映射规则。
     * - Key: 源数据项名称 (例如: "[SN] (序列号)", "电池电压")
     * - Value: 包含目标单元格地址等信息的规则对象
     */
    private Map<String, List<MappingRule>> mappingRules;

    /*
     * 从前端选中的、需要填充到报告中的日志数据记录列表。
     */
    private List<LogRecord> logData;
}