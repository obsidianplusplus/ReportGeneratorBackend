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

    private static final int CHARTS_PER_ROW = 2;
    private static final int CHART_WIDTH = 10;
    private static final int CHART_HEIGHT = 20;
    private static final int CHART_PADDING_ROWS = 2;
    private static final int CHART_PADDING_COLS = 1;

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
        if (excelFileBytes == null || excelFileBytes.length == 0) throw new IllegalArgumentException("Excel文件字节为空。");
        if (request == null || request.getSeries() == null || request.getSeries().isEmpty()) throw new IllegalArgumentException("图表定义请求无效或未定义任何系列。");

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

    private void createCombinedChartSheet(XSSFWorkbook workbook, ExcelChartRequest request) {
        String sheetName = request.getCombinedSheetName();
        if (sheetName == null || sheetName.trim().isEmpty()) {
            sheetName = "图表汇总";
        }

        XSSFSheet chartSheet = workbook.createSheet(sheetName.trim());
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        int chartCount = 0;

        for (SeriesDefinitionExcel seriesDef : request.getSeries()) {
            List<Double> dataPoints = extractDataPoints(workbook, seriesDef);
            if (dataPoints.isEmpty()) continue;

            String safeName = getSafeSheetName(seriesDef.getName());
            String uniqueSuffix = safeName + "_" + chartCount;
            XSSFSheet dataSheet = workbook.createSheet("Data_" + uniqueSuffix);
            XSSFSheet xSheet = workbook.createSheet("XAxis_" + uniqueSuffix);

            for (int i = 0; i < dataPoints.size(); i++) {
                dataSheet.createRow(i).createCell(0).setCellValue(dataPoints.get(i));
                xSheet.createRow(i).createCell(0).setCellValue(i + 1);
            }

            int rowNum = (chartCount / CHARTS_PER_ROW) * (CHART_HEIGHT + CHART_PADDING_ROWS) + CHART_PADDING_ROWS;
            int colNum = (chartCount % CHARTS_PER_ROW) * (CHART_WIDTH + CHART_PADDING_COLS) + CHART_PADDING_COLS;

            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colNum, rowNum, colNum + CHART_WIDTH, rowNum + CHART_HEIGHT);
            createChartObject(chartSheet, anchor, request, seriesDef, dataSheet, xSheet, dataPoints.size());

            workbook.setSheetHidden(workbook.getSheetIndex(dataSheet.getSheetName()), true);
            workbook.setSheetHidden(workbook.getSheetIndex(xSheet.getSheetName()), true);
            chartCount++;
        }
    }

    private void createSingleChart(XSSFWorkbook workbook, ExcelChartRequest request, SeriesDefinitionExcel seriesDef) {
        List<Double> dataPoints = extractDataPoints(workbook, seriesDef);
        if (dataPoints.isEmpty()) return;

        String safeName = getSafeSheetName(seriesDef.getName());
        XSSFSheet dataSheet = workbook.createSheet("Data_" + safeName);
        XSSFSheet xSheet = workbook.createSheet("XAxis_" + safeName);

        for (int i = 0; i < dataPoints.size(); i++) {
            dataSheet.createRow(i).createCell(0).setCellValue(dataPoints.get(i));
            xSheet.createRow(i).createCell(0).setCellValue(i + 1);
        }

        XSSFSheet chartSheet = workbook.createSheet("Chart_" + safeName);
        XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 2, 15, 32);

        createChartObject(chartSheet, anchor, request, seriesDef, dataSheet, xSheet, dataPoints.size());

        workbook.setSheetHidden(workbook.getSheetIndex(dataSheet.getSheetName()), true);
        workbook.setSheetHidden(workbook.getSheetIndex(xSheet.getSheetName()), true);
    }

    private List<Double> extractDataPoints(XSSFWorkbook workbook, SeriesDefinitionExcel seriesDef) {
        List<Double> dataPoints = new ArrayList<>();
        String sourceSheetName = seriesDef.getSheetName();
        if (sourceSheetName == null || sourceSheetName.trim().isEmpty()) {
            System.err.println("警告: 系列 '" + seriesDef.getName() + "' 未指定有效的工作表名称(sheetName)。");
            return dataPoints;
        }

        Sheet sourceSheet = workbook.getSheet(sourceSheetName);
        if (sourceSheet == null) {
            System.err.println("警告: 在工作簿中未找到名为 '" + sourceSheetName + "' 的工作表。");
            return dataPoints;
        }

        if (seriesDef.getDataAddresses() != null) {
            for (String address : seriesDef.getDataAddresses()) {
                String[] parts = address.split("_");
                int col = Integer.parseInt(parts[0]);
                int row = Integer.parseInt(parts[1]);
                Row sourceRow = sourceSheet.getRow(row);
                if (sourceRow != null) {
                    extractNumericValue(sourceRow.getCell(col)).ifPresent(dataPoints::add);
                }
            }
        }
        return dataPoints;
    }

    private void createChartObject(XSSFSheet chartSheet, XSSFClientAnchor anchor, ExcelChartRequest request,
                                   SeriesDefinitionExcel seriesDef, XSSFSheet dataSheet, XSSFSheet xSheet, int dataPointsCount) {
        XSSFChart chart = chartSheet.createDrawingPatriarch().createChart(anchor);

        String chartTitle = request.getTitle().replace("${seriesName}", seriesDef.getName());
        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);

        if (!"separate".equalsIgnoreCase(request.getOutputMode())) {
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
        }

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
                List<DetailedItem> allItems = recordsForThisSn.stream()
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
                fillDataForRecord(newSheet, request.getMappingRules(), mergedRecord, 0);
            }
            outputWorkbook.write(baos);
            return baos.toByteArray();
        }
    }

    private void fillDataForRecord(Sheet sheet, Map<String, SingleCellMapping> mappingRules, LogRecord record, int recordIndex) {
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