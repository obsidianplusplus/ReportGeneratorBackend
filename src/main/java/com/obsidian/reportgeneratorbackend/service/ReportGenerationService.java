package com.obsidian.reportgeneratorbackend.service;

import com.obsidian.reportgeneratorbackend.dto.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class ReportGenerationService {

    private static final String SN_MAPPING_KEY = "[SN] (序列号)";

    private static final Pattern NUMERIC_VALUE_PATTERN = Pattern.compile("^[\\-+]?(\\d*\\.?\\d+|\\d+\\.?\\d*)(e[+-]?\\d+)?");

    private Optional<Double> extractNumericValue(Cell cell) {
        if (cell == null) {
            return Optional.empty();
        }

        if (cell.getCellType() == CellType.NUMERIC) {
            return Optional.of(cell.getNumericCellValue());
        }

        if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().trim();
            Matcher matcher = NUMERIC_VALUE_PATTERN.matcher(cellValue);
            if (matcher.find()) {
                try {
                    String numericPart = matcher.group();
                    return Optional.of(Double.parseDouble(numericPart));
                } catch (NumberFormatException e) {
                    return Optional.empty();
                }
            }
        }

        return Optional.empty();
    }


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

    public byte[] generateChartInExcel(byte[] excelFileBytes, ExcelChartRequest request) throws IOException {
        if (excelFileBytes == null || excelFileBytes.length == 0) {
            throw new IllegalArgumentException("Excel文件字节为空。");
        }
        if (request == null || request.getSeries() == null || request.getSeries().isEmpty()) {
            throw new IllegalArgumentException("图表定义请求无效或未定义任何系列。");
        }

        try (XSSFWorkbook workbook = PoiHelper.createWorkbookFromTemplate(excelFileBytes);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            addChartSheetFromSource(workbook, request);
            workbook.write(baos);
            return baos.toByteArray();
        }
    }


    private void addChartSheetFromSource(XSSFWorkbook workbook, ExcelChartRequest request) {
        XSSFSheet dataSheet = workbook.createSheet("ChartDataSource");
        int maxDataPoints = 0;

        for (int i = 0; i < request.getSeries().size(); i++) {
            SeriesDefinitionExcel seriesDef = request.getSeries().get(i);

            Row headerRow = dataSheet.getRow(0);
            if (headerRow == null) {
                headerRow = dataSheet.createRow(0);
            }
            headerRow.createCell(i).setCellValue(seriesDef.getName());

            List<String> addresses = seriesDef.getDataAddresses();
            if (addresses.size() > maxDataPoints) {
                maxDataPoints = addresses.size();
            }

            // ======================= 核心修正点在这里 =======================
            // 使用前端传来的工作表名称来获取正确的工作表
            Sheet sourceSheet = workbook.getSheet(seriesDef.getSheetName());
            if (sourceSheet == null) {
                System.err.println("致命错误: 找不到名为 '" + seriesDef.getSheetName() + "' 的工作表。跳过此系列。");
                continue; // 跳过这个系列的处理
            }
            // =============================================================

            for (int j = 0; j < addresses.size(); j++) {
                String address = addresses.get(j);
                String[] parts = address.split("_");
                int col = Integer.parseInt(parts[0]);
                int row = Integer.parseInt(parts[1]);

                Row sourceRow = sourceSheet.getRow(row);
                if (sourceRow != null) {
                    Cell sourceCell = sourceRow.getCell(col);
                    Optional<Double> numericValue = extractNumericValue(sourceCell);

                    if (numericValue.isPresent()) {
                        Row dataRow = dataSheet.getRow(j + 1);
                        if (dataRow == null) {
                            dataRow = dataSheet.createRow(j + 1);
                        }
                        dataRow.createCell(i).setCellValue(numericValue.get());
                    }
                }
            }
        }

        if (maxDataPoints == 0) {
            workbook.removeSheetAt(workbook.getSheetIndex("ChartDataSource"));
            return;
        }

        XSSFSheet xSheet = workbook.createSheet("XAxisSource");
        for (int i = 0; i < maxDataPoints; i++) {
            xSheet.createRow(i).createCell(0).setCellValue(i + 1);
        }

        XSSFSheet chartSheet = workbook.createSheet(request.getTitle());
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 2, 20, 30);
        XSSFChart chart = drawing.createChart(anchor);

        chart.setTitleText(request.getTitle());
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(request.getXAxisTitle());
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(request.getYAxisTitle());

        XDDFChartData data;
        if ("line".equalsIgnoreCase(request.getChartType())) {
            data = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
        } else {
            data = chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
        }

        XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(xSheet, new CellRangeAddress(0, maxDataPoints - 1, 0, 0));

        for (int i = 0; i < request.getSeries().size(); i++) {
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(dataSheet, new CellRangeAddress(1, maxDataPoints, i, i));

            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle(request.getSeries().get(i).getName(), null);

            if(data instanceof XDDFScatterChartData) {
                ((XDDFScatterChartData.Series)series).setMarkerStyle(MarkerStyle.CIRCLE);
                ((XDDFScatterChartData.Series)series).setSmooth(false);
            }
            if(data instanceof XDDFLineChartData) {
                ((XDDFLineChartData.Series)series).setSmooth(false);
            }
        }

        chart.plot(data);

        workbook.setSheetHidden(workbook.getSheetIndex(dataSheet), true);
        workbook.setSheetHidden(workbook.getSheetIndex(xSheet), true);
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

    private void fillDataForRecord(Sheet sheet, Map<String, SingleCellMapping> mappingRules, LogRecord record, int recordIndex) {
        mappingRules.forEach((address, cellMapping) -> {
            String[] addressParts = address.split("_");
            if (addressParts.length != 2) {
                System.err.println("警告: 无效的映射地址格式 '" + address + "'。");
                return;
            }

            int baseRow, baseCol;
            try {
                baseRow = Integer.parseInt(addressParts[0]);
                baseCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                System.err.println("警告: 映射地址中的行列索引不是有效的数字 '" + address + "'。");
                return;
            }

            List<String> formattedValues = new ArrayList<>();

            if (cellMapping != null && cellMapping.getSources() != null) {
                cellMapping.getSources().forEach(sourceRule -> {
                    findValueForKey(sourceRule.getSourceKey(), record)
                            .ifPresent(rawValue -> {
                                String formattedValue = PoiHelper.formatValue(
                                        rawValue,
                                        sourceRule.getDecimals(),
                                        sourceRule.getUnit()
                                );
                                formattedValues.add(formattedValue);
                            });
                });
            }

            if (!formattedValues.isEmpty()) {
                String finalCellValue = String.join("/", formattedValues);
                int targetCol = baseCol + recordIndex;
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