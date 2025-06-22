package com.obsidian.reportgeneratorbackend.service;

import com.obsidian.reportgeneratorbackend.dto.LogRecord;
import com.obsidian.reportgeneratorbackend.dto.MappingRule;
import com.obsidian.reportgeneratorbackend.dto.ReportGenerationRequest;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/*
 * 描述: 报告生成的核心服务类。
 *       它作为“哑处理器”，严格按照前端提供的指令执行操作。
 */
@Service
public class ReportGenerationService {

    // 特殊的SN映射键，与前端保持一致
    private static final String SN_MAPPING_KEY = "[SN] (序列号)";

    /*
     * 主入口方法，根据不同的导出模式调用相应的处理函数。
     * @param request       包含所有指令的请求对象
     * @param templateBytes Excel模板文件的原始字节
     * @return 生成的报告文件（单个Excel或Zip压缩包）的字节数组
     */
    public byte[] generateReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        switch (request.getExportMode()) {
            case SINGLE_SHEET:
                return generateSingleSheetReport(request, templateBytes);
            case ZIP_FILES:
                return generateZipFilesReport(request, templateBytes);
            case MULTI_SHEET:
                // MULTI_SHEET模式作为可选实现，逻辑与ZIP类似，但操作的是同一个Workbook的不同Sheet
                // 这里暂时返回一个错误或未实现的信息
                throw new UnsupportedOperationException("多工作簿(Multi-sheet)模式当前未实现。");
            default:
                throw new IllegalArgumentException("未知的导出模式: " + request.getExportMode());
        }
    }

    /*
     * 生成“单表合并”模式的报告。
     * 所有数据都填充到模板的第一个工作表中。
     */
    private byte[] generateSingleSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        // *** 修改点：直接调用 PoiHelper，而不是使用错误的完整路径 ***
        try (XSSFWorkbook workbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0); // 假设所有操作都在第一个sheet上

            for (int i = 0; i < request.getLogData().size(); i++) {
                LogRecord record = request.getLogData().get(i);
                fillDataForRecord(sheet, request.getMappingRules(), record, i);
            }

            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    /*
     * 生成“多文件ZIP包”模式的报告。
     * 每个日志记录都会生成一个独立的、基于模板填充的Excel文件，然后将它们压缩成一个ZIP文件。
     */
    private byte[] generateZipFilesReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        try (ByteArrayOutputStream zipBaos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(zipBaos)) {

            for (LogRecord record : request.getLogData()) {
                // 为每条记录创建一个独立的Excel文件
                // *** 修改点：直接调用 PoiHelper ***
                try (XSSFWorkbook singleRecordWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
                     ByteArrayOutputStream singleExcelBaos = new ByteArrayOutputStream()) {

                    Sheet sheet = singleRecordWorkbook.getSheetAt(0);
                    // 在此模式下，每条记录都是一个新文件，所以偏移量总是0
                    fillDataForRecord(sheet, request.getMappingRules(), record, 0);

                    singleRecordWorkbook.write(singleExcelBaos);

                    // 创建Zip条目并写入数据
                    ZipEntry zipEntry = new ZipEntry("Report_" + record.getSn() + ".xlsx");
                    zos.putNextEntry(zipEntry);
                    zos.write(singleExcelBaos.toByteArray());
                    zos.closeEntry();
                }
            }
            // 必须在返回前关闭 ZipOutputStream
            zos.finish();
            zos.close();

            return zipBaos.toByteArray();
        }
    }

    /*
     * 核心的数据填充逻辑。
     * @param sheet       目标工作表
     * @param mappingRules 映射规则
     * @param record      当前要处理的日志记录
     * @param recordIndex 记录的索引，用于计算列偏移
     */
    private void fillDataForRecord(Sheet sheet, Map<String, MappingRule> mappingRules, LogRecord record, int recordIndex) {
        // 遍历所有映射规则
        mappingRules.forEach((sourceKey, rule) -> {
            String[] addressParts = rule.getAddress().split("_");
            int baseRow = Integer.parseInt(addressParts[0]);
            int baseCol = Integer.parseInt(addressParts[1]);

            // 根据规则的源Key找到要填充的数据
            Optional<String> valueToFillOpt = findValueForKey(sourceKey, record);

            valueToFillOpt.ifPresent(valueToFill -> {
                // *** 修改点：直接调用 PoiHelper ***
                String formattedValue = PoiHelper.formatValue(valueToFill, rule.getDecimals(), rule.getUnit());

                // 计算目标单元格，通常是列进行偏移
                int targetRow = baseRow;
                int targetCol = baseCol + recordIndex; // 列偏移

                // *** 修改点：直接调用 PoiHelper ***
                PoiHelper.setCellValue(sheet, targetRow, targetCol, formattedValue);
            });
        });
    }

    /*
     * 根据源Key从日志记录中查找对应的值。
     * @param sourceKey 源Key，可能是SN或某个itemName
     * @param record    日志记录
     * @return 查找到的值，封装在Optional中
     */
    private Optional<String> findValueForKey(String sourceKey, LogRecord record) {
        // 首先检查是否是特殊的SN键
        if (SN_MAPPING_KEY.equals(sourceKey)) {
            return Optional.ofNullable(record.getSn());
        }
        // 否则，在详细测试项中查找
        if (record.getDetailedItems() == null) {
            return Optional.empty();
        }
        return record.getDetailedItems().stream()
                .filter(item -> sourceKey.equals(item.getItemName()))
                .map(item -> item.getActualValue())
                .findFirst();
    }
}