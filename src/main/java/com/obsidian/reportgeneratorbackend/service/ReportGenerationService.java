package com.obsidian.reportgeneratorbackend.service;

import com.obsidian.reportgeneratorbackend.dto.LogRecord;
import com.obsidian.reportgeneratorbackend.dto.MappingRule;
import com.obsidian.reportgeneratorbackend.dto.ReportGenerationRequest;
import org.apache.poi.ss.usermodel.*; // 引入通配符以便使用CellType等
import org.apache.poi.ss.util.CellRangeAddress; // 用于复制合并区域
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
     * @throws IOException 如果在处理文件时发生IO错误
     * @throws IllegalArgumentException 如果请求参数无效
     * @throws UnsupportedOperationException 如果请求了未实现的模式
     */
    public byte[] generateReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        // 校验必要参数
        if (request == null || request.getLogData() == null || request.getMappingRules() == null) {
            throw new IllegalArgumentException("报告生成请求数据无效。");
        }
        if (templateBytes == null || templateBytes.length == 0) {
            // 如果是ZIP模式，模板字节可能为空，但ZIP模式通常也需要模板结构
            // 这里我们强制要求有模板，因为所有模式都基于模板填充
            // 如果ZIP模式不需要模板，这块逻辑需要调整
            throw new IllegalArgumentException("Excel模板文件字节为空。");
        }


        switch (request.getExportMode()) {
            case SINGLE_SHEET:
                // 调用已有的单表模式生成方法
                return generateSingleSheetReport(request, templateBytes);
            case ZIP_FILES:
                // 调用已有的ZIP文件模式生成方法
                return generateZipFilesReport(request, templateBytes);
            case MULTI_SHEET:
                // 调用新增的多工作簿模式生成方法
                return generateMultiSheetReport(request, templateBytes);
            default:
                throw new IllegalArgumentException("未知的导出模式: " + request.getExportMode());
        }
    }

    /*
     * 生成“单表合并”模式的报告。
     * 所有数据都填充到模板的第一个工作表中。
     */
    private byte[] generateSingleSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        // 使用 try-with-resources 确保 Workbook 和 OutputStream 被关闭
        try (XSSFWorkbook workbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            // 获取模板的第一个 Sheet，数据将填充到此 Sheet 中
            Sheet sheet = workbook.getSheetAt(0);

            // 遍历前端提供的每条日志记录
            for (int i = 0; i < request.getLogData().size(); i++) {
                LogRecord record = request.getLogData().get(i);
                // 调用核心填充方法，传入记录索引作为列偏移
                fillDataForRecord(sheet, request.getMappingRules(), record, i);
            }

            // 将填充好数据的 Workbook 写入输出流
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    /*
     * 生成“多文件ZIP包”模式的报告。
     * 每个日志记录都会生成一个独立的、基于模板填充的Excel文件，然后将它们压缩成一个ZIP文件。
     */
    private byte[] generateZipFilesReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        // 使用 try-with-resources 确保 ZipOutputStream 和其底层的 OutputStream 被关闭
        try (ByteArrayOutputStream zipBaos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(zipBaos)) {

            // 遍历前端提供的每条日志记录
            for (LogRecord record : request.getLogData()) {
                // 为每条记录创建一个独立的 Excel 文件
                // 这里的逻辑是基于模板创建一个新的 Workbook，填充单条记录数据，然后写入 ZIP
                try (XSSFWorkbook singleRecordWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
                     ByteArrayOutputStream singleExcelBaos = new ByteArrayOutputStream()) {

                    // 获取新 Workbook 的第一个 Sheet
                    Sheet sheet = singleRecordWorkbook.getSheetAt(0);
                    // 调用核心填充方法，ZIP模式下每条记录单独一个文件，没有列偏移，索引为0
                    fillDataForRecord(sheet, request.getMappingRules(), record, 0);

                    // 将填充好数据的 Workbook 写入临时的 ByteArrayOutputStream
                    singleRecordWorkbook.write(singleExcelBaos);

                    // 创建 Zip 条目，使用 SN 作为文件名，并确保扩展名为 .xlsx
                    String entryName = Optional.ofNullable(record.getSn()).orElse("UnknownSN") + ".xlsx";
                    ZipEntry zipEntry = new ZipEntry(entryName);
                    zos.putNextEntry(zipEntry);
                    // 将临时的 ByteArrayOutputStream 内容写入 Zip 条目
                    zos.write(singleExcelBaos.toByteArray());
                    zos.closeEntry(); // 关闭当前 Zip 条目
                } // singleRecordWorkbook 和 singleExcelBaos 在此关闭
            }
            // 必须在返回前调用 finish() 完成 Zip 文件的构建，并关闭 ZipOutputStream
            zos.finish();
            zos.close();

            // 返回完整的 Zip 文件字节数组
            return zipBaos.toByteArray();
        } // zipBaos 和 zos 在此关闭
    }

    /*
     * =================================================================
     *  ************ 新增的多工作簿模式实现 ************
     * =================================================================
     * 生成“多工作簿”模式的报告（实则为单个Excel文件内的多个Sheet）。
     * 为每条日志记录在新的Workbook中创建一个基于模板Sheet的副本，并填充数据。
     */
    private byte[] generateMultiSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        // 使用 try-with-resources 确保所有资源被关闭
        try (XSSFWorkbook templateWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes); // 加载模板Workbook
             XSSFWorkbook outputWorkbook = new XSSFWorkbook(); // 创建用于输出的新Workbook
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) { // 用于收集最终字节流

            // 获取模板Workbook的第一个Sheet作为复制源
            Sheet templateSheet = templateWorkbook.getSheetAt(0);
            if (templateSheet == null) {
                throw new IOException("模板文件不包含任何工作表。");
            }

            // 遍历前端提供的每条日志记录
            for (LogRecord record : request.getLogData()) {
                // 使用 SN 作为新 Sheet 的名称，处理null值，POI会自动处理重复名称
                String sheetName = Optional.ofNullable(record.getSn()).orElse("UnknownSN");

                // 在输出Workbook中创建一个新的Sheet
                Sheet newSheet = outputWorkbook.createSheet(sheetName);

                // *** 关键步骤：将模板Sheet的内容复制到这个新的Sheet中 ***
                copySheetContent(templateSheet, newSheet, outputWorkbook);

                // *** 关键步骤：将当前日志记录的数据填充到这个新的（已复制模板结构的）Sheet中 ***
                // 注意：这里调用的是 fillDataForRecordMultiSheet，它不应用列偏移
                fillDataForRecordMultiSheet(newSheet, request.getMappingRules(), record);
            }

            // 将包含所有填充好Sheet的输出Workbook写入输出流
            outputWorkbook.write(baos);
            return baos.toByteArray();
        } // 所有 try-with-resources 资源在此关闭
    }


    /*
     * =================================================================
     *  ************ 辅助方法 - 复制Sheet内容 (修正版) ************
     * =================================================================
     * 描述: 将源工作表的内容（包括单元格、样式、合并区域、列宽）复制到目标工作表。
     *       注意：这仅复制可见内容和基本格式，不处理复杂的图表、图片、VBA等。
     * @param sourceSheet    源工作表对象
     * @param targetSheet    目标工作表对象
     * @param targetWorkbook 目标工作簿对象 (用于在目标Workbook中创建新样式)
     */
    private void copySheetContent(Sheet sourceSheet, Sheet targetSheet, Workbook targetWorkbook) {
        // 复制列宽
        // POI的列宽是按1/256个字符宽度计算的，默认列宽是8个字符宽
        // 这里的遍历需要小心，getLastCellNum() 是 Row 的方法，这里应该遍历Sheet的列
        // 可以尝试复制所有非空行的最大列数，或者更简单的，直接复制那些设置了自定义宽度的列
        // 采用复制设置了自定义宽度的列，同时复制默认列宽
        int maxCol = 0;
        for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                // getLastCellNum() 返回最后一个单元格的列索引 + 1
                if (sourceRow.getLastCellNum() > maxCol) {
                    maxCol = sourceRow.getLastCellNum();
                }
            }
        }

        for (int i = 0; i < maxCol; i++) {
            try {
                // 检查是否设置了自定义列宽，如果设置了则复制
                // getColumnWidth() 返回的单位已经是 1/256 个字符
                if (sourceSheet.getColumnWidth(i) != sourceSheet.getDefaultColumnWidth() * 256) {
                    targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
                }
            } catch (Exception e) {
                // 有时候获取列宽会抛异常，忽略之
                // System.err.println("Warning: Failed to copy column width for column " + i + ": " + e.getMessage());
            }
        }
        // 复制默认列宽
        targetSheet.setDefaultColumnWidth(sourceSheet.getDefaultColumnWidth());


        // 复制合并区域
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            // 直接添加到目标 Sheet 中，区域定义是相对的
            targetSheet.addMergedRegion(mergedRegion);
        }

        // 遍历源工作表的行
        for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                // 在目标 Sheet 中创建对应的行
                Row targetRow = targetSheet.createRow(i);
                // 复制行高
                targetRow.setHeight(sourceRow.getHeight());

                // 遍历源行中的单元格
                // *** 修正点：这里使用 sourceRow.getLastCellNum() ***
                // getLastCellNum() 返回最后一个单元格索引 + 1
                for (int j = sourceRow.getFirstCellNum(); j < sourceRow.getLastCellNum(); j++) {
                    // getCell(j) 返回 null 如果单元格不存在，这是期望的行为
                    Cell sourceCell = sourceRow.getCell(j);
                    if (sourceCell != null) {
                        // 在目标行中创建对应的单元格，复制类型
                        Cell targetCell = targetRow.createCell(j, sourceCell.getCellType());

                        // 复制单元格值 (根据单元格类型)
                        switch (sourceCell.getCellType()) {
                            case STRING:
                                targetCell.setCellValue(sourceCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                // 复制数值，注意日期/时间格式需要单独处理，此处简化只复制数值
                                targetCell.setCellValue(sourceCell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                // 复制公式，但要注意相对引用的行为可能与预期不同
                                // 尝试复制公式，失败则尝试复制当前计算出的数值
                                try {
                                    targetCell.setCellFormula(sourceCell.getCellFormula());
                                } catch (Exception e) {
                                    // 复制公式失败，可能由于复杂引用或外部链接，尝试复制值
                                    try {
                                        if (sourceCell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            targetCell.setCellValue(sourceCell.getNumericCellValue());
                                        } else if (sourceCell.getCachedFormulaResultType() == CellType.STRING) {
                                            targetCell.setCellValue(sourceCell.getStringCellValue());
                                        } // 其他类型忽略或按默认处理
                                    } catch (Exception ignore) {
                                        // 再次失败则忽略
                                    }
                                }
                                break;
                            case BLANK:
                                // 空白单元格无需特别处理，createCell() 默认就是空白
                                break;
                            case ERROR:
                                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                                break;
                            default:
                                // 处理其他类型或忽略
                                break;
                        }

                        // 复制单元格样式
                        // 注意：不能直接使用源Workbook的CellStyle对象到目标Workbook，必须克隆
                        CellStyle sourceStyle = sourceCell.getCellStyle();
                        // 查找或创建一个匹配源样式的CellStyle，避免创建过多重复样式
                        CellStyle targetStyle = targetWorkbook.createCellStyle();
                        targetStyle.cloneStyleFrom(sourceStyle); // 克隆样式属性
                        targetCell.setCellStyle(targetStyle);

                        // 复制评论、超链接等（如果需要更完整的复制，这部分相对复杂，此处仅处理基本内容和样式）
                        // if (sourceCell.getCellComment() != null) { ... }
                        // if (sourceCell.getHyperlink() != null) { ... }
                    }
                } // 单元格遍历结束
            } // 行非空检查结束
        } // 行遍历结束
    }

    /*
     * =================================================================
     *  ************ 辅助方法 - 填充单条记录数据到Sheet (多Sheet模式专用) ************
     * =================================================================
     * 描述: 将单条日志记录的数据根据映射规则填充到指定的Sheet中。
     *       此方法用于多工作簿模式，不应用记录索引的列偏移。
     * @param sheet       目标工作表
     * @param mappingRules 映射规则
     * @param record      当前要处理的日志记录
     */
    private void fillDataForRecordMultiSheet(Sheet sheet, Map<String, MappingRule> mappingRules, LogRecord record) {
        // 遍历前端提供的所有映射规则
        mappingRules.forEach((sourceKey, rule) -> {
            // 解析映射规则中的目标单元格地址 (格式为 "行_列")
            String[] addressParts = rule.getAddress().split("_");
            if (addressParts.length != 2) {
                // 如果地址格式错误，则跳过此规则并打印警告
                System.err.println("警告: 无效的映射地址格式 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }

            // 将解析出的行和列索引转换为整数 (POI是0-based索引)
            int targetRow, targetCol;
            try {
                targetRow = Integer.parseInt(addressParts[0]);
                targetCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                System.err.println("警告: 映射地址中的行列索引不是有效的数字 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }


            // 根据规则的源Key ([SN] 或 测试项名称) 从当前日志记录中查找对应的值
            Optional<String> valueToFillOpt = findValueForKey(sourceKey, record);

            // 如果找到了对应的值
            valueToFillOpt.ifPresent(valueToFill -> {
                // 使用 PoiHelper 辅助类格式化值（应用小数位数和单位）
                String formattedValue = PoiHelper.formatValue(valueToFill, rule.getDecimals(), rule.getUnit());
                // 调用 PoiHelper 辅助类安全地将格式化后的值写入目标 Sheet 的指定单元格
                PoiHelper.setCellValue(sheet, targetRow, targetCol, formattedValue);
            });
            // 如果未找到对应的值，则不做任何操作（单元格保持模板的原始内容或为空）
        });
    }

    /*
     * 描述: 根据源Key从日志记录中查找对应的值。
     * @param sourceKey 源Key，可能是SN或某个itemName
     * @param record    日志记录
     * @return 查找到的值，封装在Optional中。如果未找到或对应值为null/空，则返回Optional.empty()。
     */
    private Optional<String> findValueForKey(String sourceKey, LogRecord record) {
        // 首先检查是否是特殊的SN键
        if (SN_MAPPING_KEY.equals(sourceKey)) {
            // 返回SN的值，使用 Optional.ofNullable 处理SN可能为null的情况
            return Optional.ofNullable(record.getSn());
        }
        // 否则，在详细测试项列表中查找匹配的 itemName
        // 检查 detailedItems 是否为 null 以避免 NullPointerException
        if (record.getDetailedItems() == null) {
            return Optional.empty();
        }
        return record.getDetailedItems().stream()
                .filter(item -> sourceKey.equals(item.getItemName())) // 找到 itemName 匹配的详细测试项
                .map(item -> item.getActualValue()) // 获取该测试项的实际值
                .findFirst(); // 返回找到的第一个实际值，如果没有匹配项则返回 Optional.empty()
    }

    // 核心的数据填充逻辑 (用于 SINGLE_SHEET 和 ZIP_FILES 模式)
    // 这个方法保持不变，因为它需要根据记录索引进行列偏移
    /*
     * 描述: 核心的数据填充逻辑，将单条日志记录的数据根据映射规则填充到指定Sheet中，并应用记录索引的列偏移。
     *       此方法用于单表模式和ZIP文件模式。
     * @param sheet       目标工作表
     * @param mappingRules 映射规则
     * @param record      当前要处理的日志记录
     * @param recordIndex 记录的索引，用于计算目标单元格的列偏移
     */
    private void fillDataForRecord(Sheet sheet, Map<String, MappingRule> mappingRules, LogRecord record, int recordIndex) {
        // 遍历前端提供的所有映射规则
        mappingRules.forEach((sourceKey, rule) -> {
            // 解析映射规则中的目标单元格地址 (格式为 "行_列")
            String[] addressParts = rule.getAddress().split("_");
            if (addressParts.length != 2) {
                System.err.println("警告: 无效的映射地址格式 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }

            // 将解析出的行和列索引转换为整数 (POI是0-based索引)
            int baseRow, baseCol;
            try {
                baseRow = Integer.parseInt(addressParts[0]);
                baseCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                System.err.println("警告: 映射地址中的行列索引不是有效的数字 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }


            // 根据规则的源Key ([SN] 或 测试项名称) 从当前日志记录中查找对应的值
            Optional<String> valueToFillOpt = findValueForKey(sourceKey, record);

            // 如果找到了对应的值
            valueToFillOpt.ifPresent(valueToFill -> {
                // 使用 PoiHelper 辅助类格式化值（应用小数位数和单位）
                String formattedValue = PoiHelper.formatValue(valueToFill, rule.getDecimals(), rule.getUnit());

                // 计算最终的目标单元格位置：行保持不变，列根据记录索引进行偏移
                int targetRow = baseRow;
                int targetCol = baseCol + recordIndex; // *** 这里的列偏移是关键，区分于 fillDataForRecordMultiSheet ***

                // 调用 PoiHelper 辅助类安全地将格式化后的值写入目标 Sheet 的指定单元格
                PoiHelper.setCellValue(sheet, targetRow, targetCol, formattedValue);
            });
            // 如果未找到对应的值，则不做任何操作
        });
    }

}