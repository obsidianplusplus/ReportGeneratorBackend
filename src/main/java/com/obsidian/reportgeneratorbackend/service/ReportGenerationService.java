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

    // ======================= 新增：图表布局常量 =======================
    private static final int CHARTS_PER_ROW = 2; // 每行排列多少个图表
    private static final int CHART_WIDTH = 10;   // 每个图表占用的列数
    private static final int CHART_HEIGHT = 20;  // 每个图表占用的行数
    private static final int CHART_PADDING_ROWS = 2; // 图表之间的垂直间距（行）
    private static final int CHART_PADDING_COLS = 1; // 图表之间的水平间距（列）
    // =============================================================

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
            if ("separate".equalsIgnoreCase(request.getOutputMode())) {
                for (SeriesDefinitionExcel seriesDef : request.getSeries()) {
                    createSingleChart(workbook, request, seriesDef);
                }
            } else {
                createCombinedChartSheet(workbook, request);
            }
            workbook.write(baos);
            return baos.toByteArray();
        }
    }

    /**
     * 【核心重构】此方法现在在一个工作表内创建和排列多个独立的图表。
     */
    private void createCombinedChartSheet(XSSFWorkbook workbook, ExcelChartRequest request) {
        XSSFSheet chartSheet = workbook.createSheet(request.getTitle());
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();

        int chartCount = 0;

        for (SeriesDefinitionExcel seriesDef : request.getSeries()) {
            // 为每个系列创建独立的数据源
            XSSFSheet dataSheet = workbook.createSheet("Data_" + getSafeSheetName(seriesDef.getName()));
            XSSFSheet xSheet = workbook.createSheet("XAxis_" + getSafeSheetName(seriesDef.getName()));

            int dataPointsCount = populateDataSourceSheets(workbook, seriesDef, dataSheet, xSheet);

            if (dataPointsCount == 0) {
                workbook.removeSheetAt(workbook.getSheetIndex(dataSheet.getSheetName()));
                workbook.removeSheetAt(workbook.getSheetIndex(xSheet.getSheetName()));
                continue;
            }

            // 计算当前图表的位置
            int rowNum = (chartCount / CHARTS_PER_ROW) * (CHART_HEIGHT + CHART_PADDING_ROWS) + CHART_PADDING_ROWS;
            int colNum = (chartCount % CHARTS_PER_ROW) * (CHART_WIDTH + CHART_PADDING_COLS) + CHART_PADDING_COLS;

            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colNum, rowNum, colNum + CHART_WIDTH, rowNum + CHART_HEIGHT);
            XSSFChart chart = drawing.createChart(anchor);

            String chartTitle = request.getTitle().replace("${seriesName}", seriesDef.getName());
            chart.setTitleText(chartTitle);
            chart.setTitleOverlay(false);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle(request.getXAxisTitle());
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle(request.getYAxisTitle());

            XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);

            XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(xSheet, new CellRangeAddress(0, dataPointsCount - 1, 0, 0));
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(dataSheet, new CellRangeAddress(0, dataPointsCount - 1, 0, 0));

            XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(xs, ys);
            series.setTitle(seriesDef.getName(), null);
            series.setMarkerStyle(MarkerStyle.CIRCLE);
            series.setSmooth(false);

            chart.plot(data);

            // 隐藏数据源工作表
            workbook.setSheetHidden(workbook.getSheetIndex(dataSheet.getSheetName()), true);
            workbook.setSheetHidden(workbook.getSheetIndex(xSheet.getSheetName()), true);

            chartCount++;
        }
    }

    private void createSingleChart(XSSFWorkbook workbook, ExcelChartRequest request, SeriesDefinitionExcel seriesDef) {
        String safeSheetName = getSafeSheetName(seriesDef.getName());
        XSSFSheet dataSheet = workbook.createSheet("Data_" + safeSheetName);
        XSSFSheet xSheet = workbook.createSheet("XAxis_" + safeSheetName);

        int dataPointsCount = populateDataSourceSheets(workbook, seriesDef, dataSheet, xSheet);

        if (dataPointsCount == 0) {
            workbook.removeSheetAt(workbook.getSheetIndex(dataSheet.getSheetName()));
            workbook.removeSheetAt(workbook.getSheetIndex(xSheet.getSheetName()));
            return;
        }

        XSSFSheet chartSheet = workbook.createSheet("Chart_" + safeSheetName);
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 2, 15, 32);
        XSSFChart chart = drawing.createChart(anchor);

        String chartTitle = request.getTitle().replace("${seriesName}", seriesDef.getName());
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(request.getXAxisTitle());
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(request.getYAxisTitle());

        XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);

        XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(xSheet, new CellRangeAddress(0, dataPointsCount - 1, 0, 0));
        XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(dataSheet, new CellRangeAddress(0, dataPointsCount - 1, 0, 0));

        XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(xs, ys);
        series.setTitle(seriesDef.getName(), null);
        series.setMarkerStyle(MarkerStyle.CIRCLE);
        series.setSmooth(false);

        chart.plot(data);

        workbook.setSheetHidden(workbook.getSheetIndex(dataSheet.getSheetName()), true);
        workbook.setSheetHidden(workbook.getSheetIndex(xSheet.getSheetName()), true);
    }

    /**
     * 新增：抽取出的公共方法，用于填充数据源工作表
     * @return 成功提取的数据点数量
     */
    private int populateDataSourceSheets(XSSFWorkbook workbook, SeriesDefinitionExcel seriesDef, XSSFSheet dataSheet, XSSFSheet xSheet) {
        List<String> addresses = seriesDef.getDataAddresses();
        int dataPointsCount = 0;

        Sheet sourceSheet = workbook.getSheet(seriesDef.getSheetName());
        if (sourceSheet == null) return 0;

        for (String address : addresses) {
            String[] parts = address.split("_");
            int col = Integer.parseInt(parts[0]);
            int row = Integer.parseInt(parts[1]);

            Row sourceRow = sourceSheet.getRow(row);
            if (sourceRow != null) {
                Cell sourceCell = sourceRow.getCell(col);
                Optional<Double> numericValue = extractNumericValue(sourceCell);
                if (numericValue.isPresent()) {
                    dataSheet.createRow(dataPointsCount).createCell(0).setCellValue(numericValue.get());
                    xSheet.createRow(dataPointsCount).createCell(0).setCellValue(dataPointsCount + 1);
                    dataPointsCount++;
                }
            }
        }
        return dataPointsCount;
    }

    /**
     * 新增：公共方法，用于生成一个安全的、符合Excel规范的工作表名称
     */
    private String getSafeSheetName(String name) {
        String safeName = name.replaceAll("[\\\\/*?\\[\\]:]", "_");
        int maxLength = 31 - 10; // 预留 "Data_" 或 "Chart_" 和一些随机数的前缀空间
        if (safeName.length() > maxLength) {
            safeName = safeName.substring(0, maxLength);
        }
        return safeName;
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