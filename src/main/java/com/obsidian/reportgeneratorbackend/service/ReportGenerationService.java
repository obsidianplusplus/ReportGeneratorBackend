package com.obsidian.reportgeneratorbackend.service;

import com.obsidian.reportgeneratorbackend.dto.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.ThreadLocalRandom;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class ReportGenerationService {

    private static final String SN_MAPPING_KEY = "[SN] (序列号)";
    private static final Pattern NUMERIC_VALUE_PATTERN = Pattern.compile("^[\\-+]?(\\d*\\.?\\d+|\\d+\\.?\\d*)(e[+-]?\\d+)?");

    private static final int CHARTS_PER_ROW = 2;
    private static final int CHART_WIDTH = 10;
    private static final int CHART_HEIGHT = 20;
    private static final int CHART_PADDING_ROWS = 2;
    private static final int CHART_PADDING_COLS = 1;

    // ===== 日志打印辅助方法 =====
    private void log(String message) {
        System.out.println("[JAVA LOG] " + message);
    }
    private void logError(String message, Exception e) {
        System.err.println("[JAVA ERROR] " + message);
        if (e != null) {
            e.printStackTrace(System.err);
        }
    }


    private Optional<CustomSourceItem> findCustomItemByName(String name, List<CustomSourceItem> items) {
        if (items == null || name == null) {
            return Optional.empty();
        }
        return items.stream().filter(item -> name.equals(item.getName())).findFirst();
    }

    private Optional<String> generateValueFromCustomItem(CustomSourceItem item) {
        if ("static".equals(item.getType())) {
            return Optional.ofNullable(item.getValue());
        }
        if ("random".equals(item.getType())) {
            if (item.getMin() == null || item.getMax() == null || item.getDecimals() == null) {
                return Optional.empty();
            }
            double randomValue = ThreadLocalRandom.current().nextDouble(item.getMin(), item.getMax());
            String pattern = "0." + String.join("", Collections.nCopies(item.getDecimals(), "0"));
            if (item.getDecimals() == 0) pattern = "0";
            DecimalFormat df = new DecimalFormat(pattern);
            return Optional.of(df.format(randomValue));
        }
        return Optional.empty();
    }


    private Optional<Double> extractNumericValue(Cell cell) {
        if (cell == null) return Optional.empty();
        if (cell.getCellType() == CellType.NUMERIC) return Optional.of(cell.getNumericCellValue());
        if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().trim();
            Matcher matcher = NUMERIC_VALUE_PATTERN.matcher(cellValue);
            if (matcher.find()) {
                try {
                    return Optional.of(Double.parseDouble(matcher.group()));
                } catch (NumberFormatException e) {
                    return Optional.empty();
                }
            }
        }
        return Optional.empty();
    }

    public byte[] generateReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        if (request == null || (request.getLogData() == null && request.getCustomSourceItems() == null) || request.getMappingRules() == null) {
            throw new IllegalArgumentException("报告生成请求数据无效。");
        }
        if (templateBytes == null || templateBytes.length == 0) {
            throw new IllegalArgumentException("Excel模板文件字节为空。");
        }

        if (request.getLogData() == null || request.getLogData().isEmpty()) {
            request.setLogData(Collections.singletonList(new LogRecord()));
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
        log("开始 'generateChartInExcel'...");
        if (excelFileBytes == null || excelFileBytes.length == 0) throw new IllegalArgumentException("Excel文件字节为空。");
        if (request == null || request.getSeries() == null || request.getSeries().isEmpty()) throw new IllegalArgumentException("图表定义请求无效或未定义任何系列。");

        try (XSSFWorkbook workbook = PoiHelper.createWorkbookFromTemplate(excelFileBytes);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            log("请求模式: " + request.getOutputMode());
            if ("separate".equalsIgnoreCase(request.getOutputMode())) {
                for (SeriesDefinitionExcel seriesDef : request.getSeries()) {
                    log("正在为系列 '" + seriesDef.getName() + "' 创建独立图表...");
                    createSingleChart(workbook, request, seriesDef);
                }
            } else {
                log("正在创建合并图表工作表...");
                createCombinedChartSheet(workbook, request);
            }
            log("所有图表处理完毕，正在将工作簿写入字节流...");
            workbook.write(baos);
            log("写入完成，返回字节数组。");
            return baos.toByteArray();
        } catch (Exception e) {
            logError("在 generateChartInExcel 过程中发生严重错误。", e);
            throw new IOException("生成图表时发生内部服务器错误。", e);
        }
    }

    private CellStyle createBoldStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private void createCombinedChartSheet(XSSFWorkbook workbook, ExcelChartRequest request) {
        String sheetName = request.getCombinedSheetName();
        if (sheetName == null || sheetName.trim().isEmpty()) {
            sheetName = "图表汇总";
        }

        XSSFSheet chartSheet = workbook.createSheet(sheetName.trim());
        log("创建合并工作表: " + sheetName);
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        int chartCount = 0;

        CellStyle boldStyle = createBoldStyle(workbook);

        for (SeriesDefinitionExcel seriesDef : request.getSeries()) {
            log("--> 开始处理合并图表中的系列: " + seriesDef.getName());
            List<Double> dataPoints = extractDataPoints(workbook, seriesDef);
            if (dataPoints.isEmpty()) {
                log("  - 警告: 系列 '" + seriesDef.getName() + "' 的A组数据点为空，跳过此图表。");
                continue;
            }

            String safeName = getSafeSheetName(seriesDef.getName());
            String uniqueSuffix = safeName + "_" + chartCount;

            int dataPointsCountA = dataPoints.size();
            List<Double> comparisonDataPoints = extractDataPointsFromAddresses(workbook, seriesDef.getSheetName(), seriesDef.getComparisonDataAddresses());
            int dataPointsCountB = comparisonDataPoints.size();
            int maxDataPoints = Math.max(dataPointsCountA, dataPointsCountB);

            log(String.format("  - 数据点统计: A组=%d, B组=%d, X轴最大长度=%d", dataPointsCountA, dataPointsCountB, maxDataPoints));

            XSSFSheet xSheet = workbook.createSheet("XAxis_" + uniqueSuffix);
            log("  - 创建X轴临时工作表: " + xSheet.getSheetName());
            for (int i = 0; i < maxDataPoints; i++) {
                xSheet.createRow(i).createCell(0).setCellValue(i + 1);
            }

            XSSFSheet dataSheet = workbook.createSheet("Data_" + uniqueSuffix);
            log("  - 创建A组数据临时工作表: " + dataSheet.getSheetName());
            for (int i = 0; i < dataPoints.size(); i++) {
                dataSheet.createRow(i).createCell(0).setCellValue(dataPoints.get(i));
            }

            int rowNum = (chartCount / CHARTS_PER_ROW) * (CHART_HEIGHT + CHART_PADDING_ROWS) + CHART_PADDING_ROWS;
            int colNum = (chartCount % CHARTS_PER_ROW) * (CHART_WIDTH + CHART_PADDING_COLS) + CHART_PADDING_COLS;

            if (request.isShowMinMax() && !dataPoints.isEmpty()) {
                int statsLabelCol = colNum;
                int statsValueCol = colNum + 1;
                int maxValueRow = rowNum - 2;
                int minValueRow = rowNum - 1;

                if (maxValueRow >= 0) {
                    Row maxRow = chartSheet.getRow(maxValueRow) == null ? chartSheet.createRow(maxValueRow) : chartSheet.getRow(maxValueRow);
                    Cell maxLabelCell = maxRow.createCell(statsLabelCol);
                    maxLabelCell.setCellValue("最大值:");
                    maxLabelCell.setCellStyle(boldStyle);
                    maxRow.createCell(statsValueCol).setCellValue(Collections.max(dataPoints));
                }
                if (minValueRow >= 0) {
                    Row minRow = chartSheet.getRow(minValueRow) == null ? chartSheet.createRow(minValueRow) : chartSheet.getRow(minValueRow);
                    Cell minLabelCell = minRow.createCell(statsLabelCol);
                    minLabelCell.setCellValue("最小值:");
                    minLabelCell.setCellStyle(boldStyle);
                    minRow.createCell(statsValueCol).setCellValue(Collections.min(dataPoints));
                }
            }

            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colNum, rowNum, colNum + CHART_WIDTH, rowNum + CHART_HEIGHT);
            log("  - 在 (" + rowNum + "," + colNum + ") 位置创建图表锚点。");
            createChartObject(chartSheet, anchor, request, seriesDef, dataSheet, xSheet, dataPoints.size(), uniqueSuffix);

            workbook.setSheetHidden(workbook.getSheetIndex(dataSheet.getSheetName()), true);
            workbook.setSheetHidden(workbook.getSheetIndex(xSheet.getSheetName()), true);
            log("  - 已隐藏临时工作表。");
            chartCount++;
        }
    }

    private void createSingleChart(XSSFWorkbook workbook, ExcelChartRequest request, SeriesDefinitionExcel seriesDef) {
        List<Double> dataPoints = extractDataPoints(workbook, seriesDef);
        if (dataPoints.isEmpty()) {
            log("警告: 系列 '" + seriesDef.getName() + "' 的A组数据点为空，跳过独立图表创建。");
            return;
        }

        CellStyle boldStyle = createBoldStyle(workbook);

        String safeName = getSafeSheetName(seriesDef.getName());
        String uniqueSuffix = safeName;

        int dataPointsCountA = dataPoints.size();
        List<Double> comparisonDataPoints = extractDataPointsFromAddresses(workbook, seriesDef.getSheetName(), seriesDef.getComparisonDataAddresses());
        int dataPointsCountB = comparisonDataPoints.size();
        int maxDataPoints = Math.max(dataPointsCountA, dataPointsCountB);

        log(String.format("独立图表 '%s': A组=%d, B组=%d, X轴最大长度=%d", safeName, dataPointsCountA, dataPointsCountB, maxDataPoints));

        XSSFSheet xSheet = workbook.createSheet("XAxis_" + uniqueSuffix);
        log("  - 创建X轴临时工作表: " + xSheet.getSheetName());
        for (int i = 0; i < maxDataPoints; i++) {
            xSheet.createRow(i).createCell(0).setCellValue(i + 1);
        }

        XSSFSheet dataSheet = workbook.createSheet("Data_" + uniqueSuffix);
        log("  - 创建A组数据临时工作表: " + dataSheet.getSheetName());
        for (int i = 0; i < dataPoints.size(); i++) {
            dataSheet.createRow(i).createCell(0).setCellValue(dataPoints.get(i));
        }

        XSSFSheet chartSheet = workbook.createSheet("Chart_" + safeName);
        log("  - 创建图表工作表: " + chartSheet.getSheetName());
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 2, 15, 32);

        if (request.isShowMinMax() && !dataPoints.isEmpty()) {
            int anchorStartRow = anchor.getRow1();
            int statsLabelCol = anchor.getCol1();
            int statsValueCol = anchor.getCol1() + 1;

            int maxValueRow = anchorStartRow - 2;
            int minValueRow = anchorStartRow - 1;

            if (maxValueRow >= 0) {
                Row maxRow = chartSheet.createRow(maxValueRow);
                Cell maxLabelCell = maxRow.createCell(statsLabelCol);
                maxLabelCell.setCellValue("最大值:");
                maxLabelCell.setCellStyle(boldStyle);
                maxRow.createCell(statsValueCol).setCellValue(Collections.max(dataPoints));
            }
            if (minValueRow >= 0) {
                Row minRow = chartSheet.createRow(minValueRow);
                Cell minLabelCell = minRow.createCell(statsLabelCol);
                minLabelCell.setCellValue("最小值:");
                minLabelCell.setCellStyle(boldStyle);
                minRow.createCell(statsValueCol).setCellValue(Collections.min(dataPoints));
            }
        }

        createChartObject(chartSheet, anchor, request, seriesDef, dataSheet, xSheet, dataPoints.size(), uniqueSuffix);

        workbook.setSheetHidden(workbook.getSheetIndex(dataSheet.getSheetName()), true);
        workbook.setSheetHidden(workbook.getSheetIndex(xSheet.getSheetName()), true);
        log("  - 已隐藏临时工作表。");
    }

    private List<Double> extractDataPointsFromAddresses(XSSFWorkbook workbook, String sheetName, List<String> addresses) {
        List<Double> dataPoints = new ArrayList<>();
        if (sheetName == null || sheetName.trim().isEmpty() || addresses == null) {
            return dataPoints;
        }

        Sheet sourceSheet = workbook.getSheet(sheetName);
        if (sourceSheet == null) {
            logError("在工作簿中未找到名为 '" + sheetName + "' 的工作表。", null);
            return dataPoints;
        }

        for (String address : addresses) {
            String[] parts = address.split("_");
            int col = Integer.parseInt(parts[0]);
            int row = Integer.parseInt(parts[1]);
            Row sourceRow = sourceSheet.getRow(row);
            if (sourceRow != null) {
                extractNumericValue(sourceRow.getCell(col)).ifPresent(dataPoints::add);
            }
        }
        return dataPoints;
    }

    private List<Double> extractDataPoints(XSSFWorkbook workbook, SeriesDefinitionExcel seriesDef) {
        return extractDataPointsFromAddresses(workbook, seriesDef.getSheetName(), seriesDef.getDataAddresses());
    }

    private void createChartObject(XSSFSheet chartSheet, XSSFClientAnchor anchor, ExcelChartRequest request,
                                   SeriesDefinitionExcel seriesDef, XSSFSheet dataSheet, XSSFSheet xSheet, int dataPointsCount, String uniqueSuffix) {
        log("    [createChartObject] 开始为系列 '" + seriesDef.getName() + "' 创建图表对象...");
        XSSFChart chart = chartSheet.createDrawingPatriarch().createChart(anchor);

        String chartTitle = request.getTitle().replace("${seriesName}", seriesDef.getName());
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);

        boolean hasComparison = seriesDef.getComparisonDataAddresses() != null && !seriesDef.getComparisonDataAddresses().isEmpty();
        if (hasComparison || !"separate".equalsIgnoreCase(request.getOutputMode())) {
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
        }

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(request.getXAxisTitle());
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(request.getYAxisTitle());

        XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);

        if (dataPointsCount > 0) {
            log("    - 准备绘制系列 A, 数据点: " + dataPointsCount);
            XDDFDataSource<Double> xs1 = XDDFDataSourcesFactory.fromNumericCellRange(xSheet, new CellRangeAddress(0, dataPointsCount - 1, 0, 0));
            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(dataSheet, new CellRangeAddress(0, dataPointsCount - 1, 0, 0));

            XDDFScatterChartData.Series series1 = (XDDFScatterChartData.Series) data.addSeries(xs1, ys1);
            series1.setTitle(seriesDef.getName() + (hasComparison ? " (A)" : ""), null);
            series1.setMarkerStyle(MarkerStyle.CIRCLE);
            series1.setSmooth(false);
            log("    - 系列 A 绘制成功。");
        } else {
            log("    - 系列 A 数据点为0, 跳过绘制。");
        }

        if (hasComparison) {
            log("    - 发现对比数据, 准备绘制系列 B...");
            List<Double> comparisonDataPoints = extractDataPointsFromAddresses(
                    (XSSFWorkbook) chartSheet.getWorkbook(),
                    seriesDef.getSheetName(),
                    seriesDef.getComparisonDataAddresses()
            );
            int comparisonDataPointsCount = comparisonDataPoints.size();
            log("    - 系列 B 数据点: " + comparisonDataPointsCount);

            if (comparisonDataPointsCount > 0) {
                String comparisonSheetName = "Data_" + uniqueSuffix + "_B";
                XSSFSheet comparisonDataSheet = chartSheet.getWorkbook().createSheet(comparisonSheetName);
                log("    - 创建B组数据临时工作表: " + comparisonSheetName);
                for (int i = 0; i < comparisonDataPointsCount; i++) {
                    comparisonDataSheet.createRow(i).createCell(0).setCellValue(comparisonDataPoints.get(i));
                }

                XDDFDataSource<Double> xs2 = XDDFDataSourcesFactory.fromNumericCellRange(xSheet, new CellRangeAddress(0, comparisonDataPointsCount - 1, 0, 0));
                XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(comparisonDataSheet, new CellRangeAddress(0, comparisonDataPointsCount - 1, 0, 0));

                XDDFScatterChartData.Series series2 = (XDDFScatterChartData.Series) data.addSeries(xs2, ys2);
                series2.setTitle(seriesDef.getName() + " (B)", null);
                series2.setMarkerStyle(MarkerStyle.SQUARE);
                series2.setSmooth(false);

                chartSheet.getWorkbook().setSheetHidden(chartSheet.getWorkbook().getSheetIndex(comparisonDataSheet.getSheetName()), true);
                log("    - 系列 B 绘制成功，并隐藏其临时工作表。");
            } else {
                log("    - 系列 B 数据点为0, 跳过绘制。");
            }
        }

        chart.plot(data);
        log("    - 图表对象 '" + chartTitle + "' 已完成绘制。");
    }

    private String getSafeSheetName(String name) {
        String safeName = name.replaceAll("[\\\\/*?\\[\\]:]", "_").trim();
        int maxLength = 31 - 15;
        if (safeName.length() > maxLength) safeName = safeName.substring(0, maxLength);
        return safeName;
    }

    private String toExcelAddress(String underscoreAddress) {
        if (underscoreAddress == null || !underscoreAddress.contains("_")) return "";
        String[] parts = underscoreAddress.split("_");
        int c = Integer.parseInt(parts[0]);
        int r = Integer.parseInt(parts[1]);
        String colName = "";
        int dividend = c + 1;
        while (dividend > 0) {
            int modulo = (dividend - 1) % 26;
            colName = (char)(65 + modulo) + colName;
            dividend = (int)((dividend - modulo) / 26);
        }
        return colName + (r + 1);
    }

    private byte[] generateSingleSheetReport(ReportGenerationRequest request, byte[] templateBytes) throws IOException {
        try (XSSFWorkbook workbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 0; i < request.getLogData().size(); i++) {
                LogRecord record = request.getLogData().get(i);
                fillDataForRecord(sheet, request, record, i);
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
                List<DetailedItem> allItems = recordsForThisSn.stream()
                        .filter(r -> r.getDetailedItems() != null)
                        .flatMap(r -> r.getDetailedItems().stream())
                        .collect(Collectors.toList());
                mergedRecord.setDetailedItems(allItems);

                try (XSSFWorkbook singleRecordWorkbook = PoiHelper.createWorkbookFromTemplate(templateBytes);
                     ByteArrayOutputStream singleExcelBaos = new ByteArrayOutputStream()) {
                    Sheet sheet = singleRecordWorkbook.getSheetAt(0);
                    fillDataForRecord(sheet, request, mergedRecord, 0);
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
            if (templateSheet == null) throw new IOException("模板文件不包含任何工作表。");
            for (Map.Entry<String, List<LogRecord>> entry : groupedBySn.entrySet()) {
                String sn = entry.getKey();
                List<LogRecord> recordsForThisSn = entry.getValue();
                String sheetName = sn.replaceAll("[\\\\/*?\\[\\]:]", "_");
                if (sheetName.length() > 31) sheetName = sheetName.substring(0, 31);
                Sheet newSheet = outputWorkbook.createSheet(sheetName);
                copySheetContent(templateSheet, newSheet, outputWorkbook);
                LogRecord mergedRecord = new LogRecord();
                mergedRecord.setSn(sn);
                List<DetailedItem> allItems = recordsForThisSn.stream()
                        .filter(r -> r.getDetailedItems() != null)
                        .flatMap(r -> r.getDetailedItems().stream())
                        .collect(Collectors.toList());
                mergedRecord.setDetailedItems(allItems);
                fillDataForRecord(newSheet, request, mergedRecord, 0);
            }
            outputWorkbook.write(baos);
            return baos.toByteArray();
        }
    }

    private void fillDataForRecord(Sheet sheet, ReportGenerationRequest request, LogRecord record, int recordIndex) {
        Map<String, SingleCellMapping> mappingRules = request.getMappingRules();

        mappingRules.forEach((address, cellMapping) -> {
            String[] addressParts = address.split("_");
            if (addressParts.length != 2) return;
            int baseRow, baseCol;
            try {
                baseRow = Integer.parseInt(addressParts[0]);
                baseCol = Integer.parseInt(addressParts[1]);
            } catch (NumberFormatException e) {
                return;
            }
            List<String> formattedValues = new ArrayList<>();
            if (cellMapping != null && cellMapping.getSources() != null) {
                cellMapping.getSources().forEach(sourceRule -> {

                    Optional<String> maybeValue;
                    Optional<CustomSourceItem> maybeCustom = findCustomItemByName(sourceRule.getSourceKey(), request.getCustomSourceItems());

                    if (maybeCustom.isPresent()) {
                        maybeValue = generateValueFromCustomItem(maybeCustom.get());
                    } else {
                        maybeValue = findValueInLogRecord(sourceRule.getSourceKey(), record);
                    }

                    maybeValue.ifPresent(rawValue -> {
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

    private Optional<String> findValueInLogRecord(String sourceKey, LogRecord record) {
        if (SN_MAPPING_KEY.equals(sourceKey)) return Optional.ofNullable(record.getSn());
        if (record.getDetailedItems() == null) return Optional.empty();
        return record.getDetailedItems().stream()
                .filter(item -> sourceKey.equals(item.getItemName()))
                .map(DetailedItem::getActualValue)
                .findFirst();
    }

    private void copySheetContent(Sheet sourceSheet, Sheet targetSheet, Workbook targetWorkbook) {
        int maxCol = 0;
        for (int i = sourceSheet.getFirstRowNum(); i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null && sourceRow.getLastCellNum() > maxCol) maxCol = sourceRow.getLastCellNum();
        }
        for (int i = 0; i < maxCol; i++) {
            targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
        }
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            targetSheet.addMergedRegion(sourceSheet.getMergedRegion(i));
        }

        PrintSetup sourcePrintSetup = sourceSheet.getPrintSetup();
        PrintSetup targetPrintSetup = targetSheet.getPrintSetup();
        targetPrintSetup.setPaperSize(sourcePrintSetup.getPaperSize());
        targetPrintSetup.setScale(sourcePrintSetup.getScale());
        targetPrintSetup.setFitWidth(sourcePrintSetup.getFitWidth());
        targetPrintSetup.setFitHeight(sourcePrintSetup.getFitHeight());
        targetPrintSetup.setLandscape(sourcePrintSetup.getLandscape());
        targetPrintSetup.setLeftToRight(sourcePrintSetup.getLeftToRight());
        targetPrintSetup.setNoOrientation(sourcePrintSetup.getNoOrientation());
        targetPrintSetup.setUsePage(sourcePrintSetup.getUsePage());
        targetPrintSetup.setHResolution(sourcePrintSetup.getHResolution());
        targetPrintSetup.setVResolution(sourcePrintSetup.getVResolution());
        targetPrintSetup.setHeaderMargin(sourcePrintSetup.getHeaderMargin());
        targetPrintSetup.setFooterMargin(sourcePrintSetup.getFooterMargin());
        targetPrintSetup.setCopies(sourcePrintSetup.getCopies());

        targetSheet.setMargin(Sheet.TopMargin, sourceSheet.getMargin(Sheet.TopMargin));
        targetSheet.setMargin(Sheet.BottomMargin, sourceSheet.getMargin(Sheet.BottomMargin));
        targetSheet.setMargin(Sheet.LeftMargin, sourceSheet.getMargin(Sheet.LeftMargin));
        targetSheet.setMargin(Sheet.RightMargin, sourceSheet.getMargin(Sheet.RightMargin));

        if (sourceSheet.getHeader() != null) {
            targetSheet.getHeader().setCenter(sourceSheet.getHeader().getCenter());
            targetSheet.getHeader().setLeft(sourceSheet.getHeader().getLeft());
            targetSheet.getHeader().setRight(sourceSheet.getHeader().getRight());
        }
        if (sourceSheet.getFooter() != null) {
            targetSheet.getFooter().setCenter(sourceSheet.getFooter().getCenter());
            targetSheet.getFooter().setLeft(sourceSheet.getFooter().getLeft());
            targetSheet.getFooter().setRight(sourceSheet.getFooter().getRight());
        }

        if (sourceSheet instanceof XSSFSheet && targetSheet instanceof XSSFSheet) {
            XSSFSheet xssfSourceSheet = (XSSFSheet) sourceSheet;
            XSSFSheet xssfTargetSheet = (XSSFSheet) targetSheet;

            int sourceSheetIndex = sourceSheet.getWorkbook().getSheetIndex(sourceSheet);
            String printArea = sourceSheet.getWorkbook().getPrintArea(sourceSheetIndex);
            if(printArea != null) {
                targetSheet.getWorkbook().setPrintArea(
                        targetSheet.getWorkbook().getSheetIndex(targetSheet),
                        printArea.substring(printArea.indexOf('!')+1)
                );
            }

            CellRangeAddress repeatingRows = xssfSourceSheet.getRepeatingRows();
            if (repeatingRows != null) {
                xssfTargetSheet.setRepeatingRows(repeatingRows);
            }

            CellRangeAddress repeatingCols = xssfSourceSheet.getRepeatingColumns();
            if (repeatingCols != null) {
                xssfTargetSheet.setRepeatingColumns(repeatingCols);
            }
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
                            case STRING: targetCell.setCellValue(sourceCell.getStringCellValue()); break;
                            case NUMERIC: targetCell.setCellValue(sourceCell.getNumericCellValue()); break;
                            case BOOLEAN: targetCell.setCellValue(sourceCell.getBooleanCellValue()); break;
                            case FORMULA:
                                try {
                                    targetCell.setCellFormula(sourceCell.getCellFormula());
                                } catch (Exception e) {
                                    try {
                                        if (sourceCell.getCachedFormulaResultType() == CellType.NUMERIC) targetCell.setCellValue(sourceCell.getNumericCellValue());
                                        else if (sourceCell.getCachedFormulaResultType() == CellType.STRING) targetCell.setCellValue(sourceCell.getStringCellValue());
                                    } catch (Exception ignore) {}
                                }
                                break;
                            default: break;
                        }
                        CellStyle sourceStyle = sourceCell.getCellStyle();
                        CellStyle targetStyle = targetWorkbook.createCellStyle();
                        targetStyle.cloneStyleFrom(sourceStyle);
                        targetCell.setCellStyle(targetStyle);
                    }
                }
            }
        }

        if (sourceSheet.getDrawingPatriarch() instanceof XSSFDrawing) {
            XSSFDrawing sourceDrawing = (XSSFDrawing) sourceSheet.getDrawingPatriarch();
            XSSFDrawing targetDrawing = (XSSFDrawing) targetSheet.createDrawingPatriarch();
            for (XSSFShape shape : sourceDrawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture sourcePicture = (XSSFPicture) shape;
                    XSSFPictureData sourcePictureData = sourcePicture.getPictureData();
                    if (sourcePicture.getAnchor() instanceof XSSFClientAnchor) {
                        XSSFClientAnchor sourceClientAnchor = (XSSFClientAnchor) sourcePicture.getAnchor();
                        int targetPictureIndex = targetWorkbook.addPicture(sourcePictureData.getData(), sourcePictureData.getPictureType());
                        targetDrawing.createPicture(sourceClientAnchor, targetPictureIndex);
                    }
                }
            }
        }
    }
}