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
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
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
     * 生成“多文件ZIP包”模式的报告。
     * 每个日志记录都会生成一个独立的、基于模板填充的Excel文件，然后将它们压缩成一个ZIP文件。
     */
    private byte[] generateZipFilesReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        try (ByteArrayOutputStream zipBaos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(zipBaos)) {

            final Set<String> usedEntryNames = new HashSet<>();

            for (LogRecord record : request.getLogData()) {
                try (XSSFWorkbook singleRecordWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
                     ByteArrayOutputStream singleExcelBaos = new ByteArrayOutputStream()) {

                    Sheet sheet = singleRecordWorkbook.getSheetAt(0);
                    fillDataForRecord(sheet, request.getMappingRules(), record, 0);

                    singleRecordWorkbook.write(singleExcelBaos);

                    String baseName = Optional.ofNullable(record.getSn()).orElse("UnknownSN");
                    String safeBaseName = baseName.replaceAll("[\\\\/:*?\"<>|]", "_");
                    String entryName = safeBaseName + ".xlsx";
                    int counter = 1;

                    while (usedEntryNames.contains(entryName)) {
                        entryName = safeBaseName + " (" + counter++ + ").xlsx";
                    }
                    usedEntryNames.add(entryName);

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
     *
     * 【核心重构】此方法现在会先按 SN 对所有日志记录进行分组聚合。
     *  - 每个唯一的 SN 将生成一个独立的 Sheet。
     *  - 同一个 SN 下来自不同工位的数据，其测试项将被【合并】后填充到该 SN 对应的 Sheet 的【同一列】中。
     */
    private byte[] generateMultiSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        // 1. 按 SN 对日志数据进行分组聚合
        Map<String, List<LogRecord>> groupedBySn = request.getLogData().stream()
                .filter(record -> record.getSn() != null && !record.getSn().isEmpty())
                .collect(Collectors.groupingBy(LogRecord::getSn));

        try (XSSFWorkbook templateWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
             XSSFWorkbook outputWorkbook = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            Sheet templateSheet = templateWorkbook.getSheetAt(0);
            if (templateSheet == null) {
                throw new IOException("模板文件不包含任何工作表。");
            }

            // 2. 遍历按 SN 分组后的 Map
            for (Map.Entry<String, List<LogRecord>> entry : groupedBySn.entrySet()) {
                String sn = entry.getKey();
                List<LogRecord> recordsForThisSn = entry.getValue();

                // 3. 为每个唯一的 SN 创建一个 Sheet
                String sheetName = sn.replaceAll("[\\\\/*?\\[\\]:]", "_");
                if (sheetName.length() > 31) {
                    sheetName = sheetName.substring(0, 31);
                }
                Sheet newSheet = outputWorkbook.createSheet(sheetName);

                // 复制模板内容到新 Sheet
                copySheetContent(templateSheet, newSheet, outputWorkbook);

                // 4. 创建一个“合并后”的记录
                LogRecord mergedRecord = new LogRecord();
                mergedRecord.setSn(sn);

                // 5. 将该SN下的所有记录的 detailedItems 合并到一个列表中
                List<com.obsidian.reportgeneratorbackend.dto.DetailedItem> allItems = recordsForThisSn.stream()
                        .filter(r -> r.getDetailedItems() != null)
                        .flatMap(r -> r.getDetailedItems().stream())
                        .collect(Collectors.toList());
                mergedRecord.setDetailedItems(allItems);

                // 6. 调用 fillDataForRecord，但只调用一次，并且 recordIndex 永远是 0
                //    这意味着不再有列的偏移
                fillDataForRecord(newSheet, request.getMappingRules(), mergedRecord, 0);
            }

            outputWorkbook.write(baos);
            return baos.toByteArray();
        }
    }

    /*
     * 描述: 核心的数据填充逻辑，将单条日志记录的数据根据映射规则填充到指定Sheet中，并应用记录索引的列偏移。
     */
    private void fillDataForRecord(Sheet sheet, Map<String, MappingRule> mappingRules, LogRecord record, int recordIndex) {
        mappingRules.forEach((sourceKey, rule) -> {
            String[] addressParts = rule.getAddress().split("_");
            if (addressParts.length != 2) {
                System.err.println("警告: 无效的映射地址格式 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }

            int baseRow, baseCol;
            try {
                baseRow = Integer.parseInt(addressParts[0]);
                baseCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                System.err.println("警告: 映射地址中的行列索引不是有效的数字 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }

            Optional<String> valueToFillOpt = findValueForKey(sourceKey, record);
            valueToFillOpt.ifPresent(valueToFill -> {
                String formattedValue = PoiHelper.formatValue(valueToFill, rule.getDecimals(), rule.getUnit());
                int targetRow = baseRow;
                int targetCol = baseCol + recordIndex;
                PoiHelper.setCellValue(sheet, targetRow, targetCol, formattedValue);
            });
        });
    }

    /*
     * 描述: 辅助方法，将单条日志记录的数据根据映射规则填充到指定的Sheet中。
     *       此方法用于多工作簿模式（原始版本，现已废弃，由聚合逻辑取代）。
     *       保留此方法以供参考或未来可能的不同模式。
     */
    private void fillDataForRecordMultiSheet(Sheet sheet, Map<String, MappingRule> mappingRules, LogRecord record) {
        mappingRules.forEach((sourceKey, rule) -> {
            String[] addressParts = rule.getAddress().split("_");
            if (addressParts.length != 2) {
                System.err.println("警告: 无效的映射地址格式 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }

            int targetRow, targetCol;
            try {
                targetRow = Integer.parseInt(addressParts[0]);
                targetCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                System.err.println("警告: 映射地址中的行列索引不是有效的数字 '" + rule.getAddress() + "' 对于源 '" + sourceKey + "'。");
                return;
            }

            Optional<String> valueToFillOpt = findValueForKey(sourceKey, record);
            valueToFillOpt.ifPresent(valueToFill -> {
                String formattedValue = PoiHelper.formatValue(valueToFill, rule.getDecimals(), rule.getUnit());
                PoiHelper.setCellValue(sheet, targetRow, targetCol, formattedValue);
            });
        });
    }

    /*
     * 描述: 根据源Key从日志记录中查找对应的值。
     */
    private Optional<String> findValueForKey(String sourceKey, LogRecord record) {
        if (SN_MAPPING_KEY.equals(sourceKey)) {
            return Optional.ofNullable(record.getSn());
        }
        if (record.getDetailedItems() == null) {
            return Optional.empty();
        }
        return record.getDetailedItems().stream()
                .filter(item -> sourceKey.equals(item.getItemName()))
                .map(item -> item.getActualValue())
                .findFirst();
    }

    /*
     * 描述: 将源工作表的内容（包括单元格、样式、合并区域、列宽、图片）复制到目标工作表。
     */
    private void copySheetContent(Sheet sourceSheet, Sheet targetSheet, Workbook targetWorkbook) {
        int maxCol = 0;
        for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null && sourceRow.getLastCellNum() > maxCol) {
                maxCol = sourceRow.getLastCellNum();
            }
        }

        for (int i = 0; i < maxCol; i++) {
            targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
        }
        targetSheet.setDefaultColumnWidth(sourceSheet.getDefaultColumnWidth());

        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            targetSheet.addMergedRegion(sourceSheet.getMergedRegion(i));
        }

        for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                Row targetRow = targetSheet.createRow(i);
                targetRow.setHeight(sourceRow.getHeight());

                for (int j = sourceRow.getFirstCellNum(); j < sourceRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceRow.getCell(j);
                    if (sourceCell != null) {
                        Cell targetCell = targetRow.createCell(j, sourceCell.getCellType());

                        switch (sourceCell.getCellType()) {
                            case STRING:
                                targetCell.setCellValue(sourceCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                targetCell.setCellValue(sourceCell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                try {
                                    targetCell.setCellFormula(sourceCell.getCellFormula());
                                } catch (Exception e) {
                                    try {
                                        if (sourceCell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                            targetCell.setCellValue(sourceCell.getNumericCellValue());
                                        } else if (sourceCell.getCachedFormulaResultType() == CellType.STRING) {
                                            targetCell.setCellValue(sourceCell.getStringCellValue());
                                        }
                                    } catch (Exception ignore) {}
                                }
                                break;
                            case BLANK:
                                break;
                            case ERROR:
                                targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                                break;
                            default:
                                break;
                        }

                        CellStyle sourceStyle = sourceCell.getCellStyle();
                        CellStyle targetStyle = targetWorkbook.createCellStyle();
                        targetStyle.cloneStyleFrom(sourceStyle);
                        targetCell.setCellStyle(targetStyle);
                    }
                }
            }
        }

        XSSFDrawing sourceDrawing = (XSSFDrawing) sourceSheet.getDrawingPatriarch();
        if (sourceDrawing != null) {
            XSSFDrawing targetDrawing = (XSSFDrawing) targetSheet.createDrawingPatriarch();
            for (XSSFShape shape : sourceDrawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture sourcePicture = (XSSFPicture) shape;
                    XSSFPictureData sourcePictureData = sourcePicture.getPictureData();
                    if (sourcePicture.getAnchor() instanceof XSSFClientAnchor) {
                        XSSFClientAnchor sourceClientAnchor = (XSSFClientAnchor) sourcePicture.getAnchor();
                        XSSFClientAnchor targetClientAnchor = new XSSFClientAnchor(
                                sourceClientAnchor.getDx1(), sourceClientAnchor.getDy1(),
                                sourceClientAnchor.getDx2(), sourceClientAnchor.getDy2(),
                                sourceClientAnchor.getCol1(), sourceClientAnchor.getRow1(),
                                sourceClientAnchor.getCol2(), sourceClientAnchor.getRow2()
                        );
                        targetClientAnchor.setAnchorType(sourceClientAnchor.getAnchorType());
                        int targetPictureIndex = targetWorkbook.addPicture(
                                sourcePictureData.getData(),
                                sourcePictureData.getPictureType()
                        );
                        targetDrawing.createPicture(targetClientAnchor, targetPictureIndex);
                    }
                }
            }
        }
    }
}