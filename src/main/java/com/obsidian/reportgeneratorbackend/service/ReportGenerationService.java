package com.obsidian.reportgeneratorbackend.service;

import com.obsidian.reportgeneratorbackend.dto.LogRecord;
import com.obsidian.reportgeneratorbackend.dto.MappingRule;
import com.obsidian.reportgeneratorbackend.dto.ReportGenerationRequest;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Optional;
// 新增的导入
import java.util.HashSet;
import java.util.Set;
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
            throw new IllegalArgumentException("Excel模板文件字节为空。");
        }


        switch (request.getExportMode()) {
            case SINGLE_SHEET:
                return generateSingleSheetReport(request, templateBytes);
            case ZIP_FILES:
                return generateZipFilesReport(request, templateBytes);
            case MULTI_SHEET:
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
        try (XSSFWorkbook workbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 0; i < request.getLogData().size(); i++) {
                LogRecord record = request.getLogData().get(i);
                fillDataForRecord(sheet, request.getMappingRules(), record, i);
            }

            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    /*
     * =================================================================
     *  ************ 已修复和增强的ZIP文件模式实现 ************
     * =================================================================
     * 生成“多文件ZIP包”模式的报告。
     * 每个日志记录都会生成一个独立的、基于模板填充的Excel文件，然后将它们压缩成一个ZIP文件。
     *
     * 【修复】新增了对重复SN的处理逻辑，防止ZipException。
     * 【增强】新增了对SN中非法文件名字符的清理。
     */
    private byte[] generateZipFilesReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        try (ByteArrayOutputStream zipBaos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(zipBaos)) {

            // 用于跟踪已添加到ZIP中的文件名，以处理重复
            final Set<String> usedEntryNames = new HashSet<>();

            for (LogRecord record : request.getLogData()) {
                try (XSSFWorkbook singleRecordWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
                     ByteArrayOutputStream singleExcelBaos = new ByteArrayOutputStream()) {

                    Sheet sheet = singleRecordWorkbook.getSheetAt(0);
                    fillDataForRecord(sheet, request.getMappingRules(), record, 0);

                    singleRecordWorkbook.write(singleExcelBaos);

                    // --- 增强的命名逻辑 ---
                    // 1. 获取基础名称，处理SN为null的情况
                    String baseName = Optional.ofNullable(record.getSn()).orElse("UnknownSN");
                    // 2. 清理文件名中的非法字符
                    String safeBaseName = baseName.replaceAll("[\\\\/:*?\"<>|]", "_");

                    String entryName = safeBaseName + ".xlsx";
                    int counter = 1;
                    // 3. 【修复】检查文件名是否重复，如果重复则添加后缀，例如 "SN123 (1).xlsx"
                    while (usedEntryNames.contains(entryName)) {
                        entryName = safeBaseName + " (" + counter++ + ").xlsx";
                    }
                    usedEntryNames.add(entryName);
                    // --- 命名逻辑结束 ---

                    ZipEntry zipEntry = new ZipEntry(entryName);
                    zos.putNextEntry(zipEntry);
                    zos.write(singleExcelBaos.toByteArray());
                    zos.closeEntry();
                }
            }
            zos.finish();
            zos.close();

            return zipBaos.toByteArray();
        }
    }


    /*
     * 生成“多工作簿”模式的报告（实则为单个Excel文件内的多个Sheet）。
     * 为每条日志记录在新的Workbook中创建一个基于模板Sheet的副本，并填充数据。
     */
    private byte[] generateMultiSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        try (XSSFWorkbook templateWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
             XSSFWorkbook outputWorkbook = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            Sheet templateSheet = templateWorkbook.getSheetAt(0);
            if (templateSheet == null) {
                throw new IOException("模板文件不包含任何工作表。");
            }

            for (LogRecord record : request.getLogData()) {
                String sheetName = Optional.ofNullable(record.getSn()).orElse("UnknownSN");
                sheetName = sheetName.replaceAll("[\\\\/*?\\[\\]:]", "_");
                if (sheetName.length() > 31) {
                    sheetName = sheetName.substring(0, 31);
                }

                // POI的createSheet方法会自动处理重名，例如将第二个重名Sheet命名为 "SheetName (2)"
                Sheet newSheet = outputWorkbook.createSheet(sheetName);

                copySheetContent(templateSheet, newSheet, outputWorkbook);

                fillDataForRecordMultiSheet(newSheet, request.getMappingRules(), record);
            }

            outputWorkbook.write(baos);
            return baos.toByteArray();
        }
    }

    // ... copySheetContent, fillDataForRecordMultiSheet, findValueForKey, fillDataForRecord 方法保持不变 ...

    /*
     * =================================================================
     *  ************ 辅助方法 - 复制Sheet内容 (含图片复制) ************
     * =================================================================
     * 描述: 将源工作表的内容（包括单元格、样式、合并区域、列宽）复制到目标工作表。
     *       新增图片复制功能。
     *       注意：这仅复制可见内容、基本格式、合并区域、列宽和图片。
     *       不处理复杂的绘图对象（形状、文本框、图表）、评论、超链接、条件格式、数据验证等。
     * @param sourceSheet    源工作表对象
     * @param targetSheet    目标工作表对象
     * @param targetWorkbook 目标工作簿对象 (用于在目标Workbook中创建新样式和图片数据)
     */
    private void copySheetContent(Sheet sourceSheet, Sheet targetSheet, Workbook targetWorkbook) {
        // 1. 复制列宽
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


        // 2. 复制合并区域
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            // 直接添加到目标 Sheet 中，区域定义是相对的
            targetSheet.addMergedRegion(mergedRegion);
        }

        // 3. 遍历源工作表的行和单元格，复制内容和样式
        for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                // 在目标 Sheet 中创建对应的行
                Row targetRow = targetSheet.createRow(i);
                // 复制行高
                targetRow.setHeight(sourceRow.getHeight());

                // 遍历源行中的单元格
                // getLastCellNum() 返回最后一个单元格索引 + 1 (如果行没有单元格则返回-1)
                // 遍历到 < getLastCellNum() 是正确的范围
                // 如果 getLastCellNum() 是 -1，循环条件 j < -1 不会成立，这也是正确的
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
                        // 这里简化处理，直接创建新样式并克隆属性，如果性能是问题，可以实现样式缓存机制
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

        // 4. 复制图片
        XSSFDrawing sourceDrawing = (XSSFDrawing) sourceSheet.getDrawingPatriarch();
        if (sourceDrawing != null) {
            // 在目标 Sheet 中获取或创建绘图管理器
            // XSSFSheet的createDrawingPatriarch()方法返回XSSFDrawing
            XSSFDrawing targetDrawing = (XSSFDrawing) targetSheet.createDrawingPatriarch();

            // 遍历源绘图中的所有形状
            for (XSSFShape shape : sourceDrawing.getShapes()) {
                // 检查形状是否是图片
                if (shape instanceof XSSFPicture) {
                    XSSFPicture sourcePicture = (XSSFPicture) shape;
                    XSSFPictureData sourcePictureData = sourcePicture.getPictureData();
                    org.apache.poi.xssf.usermodel.XSSFAnchor sourceAnchor = sourcePicture.getAnchor();

                    // 复制锚点信息，创建一个新的 ClientAnchor
                    // 这里的锚点类型需要与源锚点匹配，ClientAnchor是最常见的
                    if (sourceAnchor instanceof XSSFClientAnchor) {
                        XSSFClientAnchor sourceClientAnchor = (XSSFClientAnchor) sourceAnchor;
                        // 创建目标 ClientAnchor，并复制属性
                        // anchor(int dx1, int dy1, int dx2, int dy2, short col1, int row1, short col2, int row2)
                        XSSFClientAnchor targetClientAnchor = new XSSFClientAnchor(
                                sourceClientAnchor.getDx1(),
                                sourceClientAnchor.getDy1(),
                                sourceClientAnchor.getDx2(),
                                sourceClientAnchor.getDy2(),
                                sourceClientAnchor.getCol1(),
                                sourceClientAnchor.getRow1(),
                                sourceClientAnchor.getCol2(),
                                sourceClientAnchor.getRow2()
                        );
                        // 复制图片比例属性
                        targetClientAnchor.setAnchorType(sourceClientAnchor.getAnchorType());


                        // 将源图片的原始字节数据添加到目标 Workbook 的图片集合中
                        int targetPictureIndex = targetWorkbook.addPicture(
                                sourcePictureData.getData(),
                                sourcePictureData.getPictureType() // 复制图片类型 (PNG, JPEG等)
                        );

                        // 在目标绘图管理器中使用新的锚点和图片索引创建新的图片
                        targetDrawing.createPicture(targetClientAnchor, targetPictureIndex);

                    } else {
                        // 如果是其他类型的锚点（例如 AbsoluteAnchor），此处不处理，可以添加警告
                        System.err.println("Warning: Ignoring non-ClientAnchor picture in source sheet.");
                    }
                }
                // TODO: 如果需要复制其他形状（文本框、简单图形等），需要在这里添加对应的 instanceof 检查和复制逻辑
                // TODO: 复制图表非常复杂，通常不在此类简单复制方法中实现
            }
        } // sourceDrawing 非空检查结束
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


            // 根据规则的源Key ([SN] 或 测试项名称) 从日志记录中查找对应的值
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