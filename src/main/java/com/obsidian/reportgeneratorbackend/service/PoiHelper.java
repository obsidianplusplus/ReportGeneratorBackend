package com.obsidian.reportgeneratorbackend.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;

/*
 * 描述: Apache POI 操作的辅助工具类。
 *       封装了单元格创建和值设定的通用逻辑，使主服务代码更简洁。
 */
public class PoiHelper {

    /*
     * 安全地设置单元格的值。
     * 如果行或单元格不存在，会自动创建它们。
     * @param sheet      工作表对象
     * @param rowIndex   行索引 (0-based)
     * @param colIndex   列索引 (0-based)
     * @param value      要设置的值
     */
    public static void setCellValue(Sheet sheet, int rowIndex, int colIndex, Object value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value != null) {
            cell.setCellValue(String.valueOf(value));
        }
    }

    /*
     * 根据规则格式化要输出的值。
     * @param actualValue 原始值
     * @param decimals    要保留的小数位数
     * @param unit        要附加的单位
     * @return 格式化后的字符串
     */
    public static String formatValue(String actualValue, Integer decimals, String unit) {
        if (actualValue == null || actualValue.trim().isEmpty()) {
            return ""; // 如果原始值为空，返回空字符串
        }

        try {
            // 尝试将值转换为数字进行格式化
            BigDecimal numberValue = new BigDecimal(actualValue);

            if (decimals != null && decimals >= 0) {
                numberValue = numberValue.setScale(decimals, RoundingMode.HALF_UP);
            }

            String formattedString = numberValue.toPlainString();

            if (unit != null && !unit.trim().isEmpty()) {
                return formattedString + " " + unit;
            } else {
                return formattedString;
            }

        } catch (NumberFormatException e) {
            // 如果无法解析为数字，则直接返回原始值，并附加单位（如果存在）
            if (unit != null && !unit.trim().isEmpty()) {
                return actualValue + " " + unit;
            }
            return actualValue;
        }
    }

    /*
     * 从字节数组模板安全地创建工作簿。
     * @param templateBytes 模板文件的字节数组
     * @return XSSFWorkbook 实例
     * @throws IOException 如果读取失败
     */
    public static XSSFWorkbook createWorkbookFromTemplate(byte[] templateBytes) throws IOException {
        try (InputStream templateStream = new ByteArrayInputStream(templateBytes)) {
            return new XSSFWorkbook(templateStream);
        }
    }
}