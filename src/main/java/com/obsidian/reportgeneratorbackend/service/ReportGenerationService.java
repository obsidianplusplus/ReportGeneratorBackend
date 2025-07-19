// src/main/java/com/obsidian/reportgeneratorbackend/service/ReportGenerationService.java
package com.obsidian.reportgeneratorbackend.service;

import com.obsidian.reportgeneratorbackend.dto.LogRecord;
import com.obsidian.reportgeneratorbackend.dto.SingleCellMapping; // <-- Import a nova classe
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
import java.util.ArrayList; // <-- Adicionar import
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

    private static final String SN_MAPPING_KEY = "[SN] (序列号)";

    public byte[] generateReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
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

    private byte[] generateZipFilesReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        Map<String, List<LogRecord>> groupedBySn = request.getLogData().stream()
                .filter(record -> record.getSn() != null && !record.getSn().isEmpty())
                .collect(Collectors.groupingBy(LogRecord::getSn));

        try (ByteArrayOutputStream zipBaos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(zipBaos)) {

            for (Map.Entry<String, List<LogRecord>> entry : groupedBySn.entrySet()) {
                String sn = entry.getKey();
                List<LogRecord> recordsForThisSn = entry.getValue();

                LogRecord mergedRecord = new LogRecord();
                mergedRecord.setSn(sn);

                List<com.obsidian.reportgeneratorbackend.dto.DetailedItem> allItems = recordsForThisSn.stream()
                        .filter(r -> r.getDetailedItems() != null)
                        .flatMap(r -> r.getDetailedItems().stream())
                        .collect(Collectors.toList());
                mergedRecord.setDetailedItems(allItems);

                try (XSSFWorkbook singleRecordWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
                     ByteArrayOutputStream singleExcelBaos = new ByteArrayOutputStream()) {

                    Sheet sheet = singleRecordWorkbook.getSheetAt(0);
                    fillDataForRecord(sheet, request.getMappingRules(), mergedRecord, 0);
                    singleRecordWorkbook.write(singleExcelBaos);

                    String safeSn = sn.replaceAll("[\\\\/:*?\"<>|]", "_");
                    ZipEntry zipEntry = new ZipEntry(safeSn + ".xlsx");
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

    private byte[] generateMultiSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
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

            for (Map.Entry<String, List<LogRecord>> entry : groupedBySn.entrySet()) {
                String sn = entry.getKey();
                List<LogRecord> recordsForThisSn = entry.getValue();

                String sheetName = sn.replaceAll("[\\\\/*?\\[\\]:]", "_");
                if (sheetName.length() > 31) {
                    sheetName = sheetName.substring(0, 31);
                }
                Sheet newSheet = outputWorkbook.createSheet(sheetName);
                copySheetContent(templateSheet, newSheet, outputWorkbook);

                LogRecord mergedRecord = new LogRecord();
                mergedRecord.setSn(sn);

                List<com.obsidian.reportgeneratorbackend.dto.DetailedItem> allItems = recordsForThisSn.stream()
                        .filter(r -> r.getDetailedItems() != null)
                        .flatMap(r -> r.getDetailedItems().stream())
                        .collect(Collectors.toList());
                mergedRecord.setDetailedItems(allItems);

                fillDataForRecord(newSheet, request.getMappingRules(), mergedRecord, 0);
            }

            outputWorkbook.write(baos);
            return baos.toByteArray();
        }
    }

    /*
     * 描述: 【V9.0 重写】核心数据填充逻辑，支持多源到一格。
     *       它遍历每个目标单元格，收集所有映射到此单元格的源的值，然后用 " / " 连接并填充。
     */
    private void fillDataForRecord(Sheet sheet, Map<String, SingleCellMapping> mappingRules, LogRecord record, int recordIndex) {
        // 遍历每个映射的目标单元格地址
        mappingRules.forEach((address, cellMapping) -> {
            // 1. 解析地址
            String[] addressParts = address.split("_");
            if (addressParts.length != 2) {
                System.err.println("警告: 无效的映射地址格式 '" + address + "'。");
                return; // continue
            }

            int baseRow, baseCol;
            try {
                baseRow = Integer.parseInt(addressParts[0]);
                baseCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                System.err.println("警告: 映射地址中的行列索引不是有效的数字 '" + address + "'。");
                return; // continue
            }

            // 2. 准备一个列表来收集该单元格所有源的格式化值
            List<String> formattedValues = new ArrayList<>();

            // 3. 遍历映射到此单元格的所有源规则
            if (cellMapping != null && cellMapping.getSources() != null) {
                cellMapping.getSources().forEach(sourceRule -> {
                    // 4. 为每个源查找其值
                    findValueForKey(sourceRule.getSourceKey(), record)
                            .ifPresent(rawValue -> {
                                // 5. 使用其独立的规则进行格式化
                                String formattedValue = PoiHelper.formatValue(
                                        rawValue,
                                        sourceRule.getDecimals(),
                                        sourceRule.getUnit()
                                );
                                formattedValues.add(formattedValue);
                            });
                });
            }

            // 6. 如果收集到了任何值，用 " / " 连接它们
            if (!formattedValues.isEmpty()) {
                String finalCellValue = String.join("/", formattedValues);

                // 7. 计算最终的目标列（考虑单表模式的偏移）
                int targetCol = baseCol + recordIndex;

                // 8. 设置单元格的值
                PoiHelper.setCellValue(sheet, baseRow, targetCol, finalCellValue);
            }
        });
    }

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