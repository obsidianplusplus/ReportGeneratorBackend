// com/obsidian/reportgeneratorbackend/service/PoiHelper.java (已修正)
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

public class PoiHelper {

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

    /**
     * 根据规则格式化要输出的值。
     * 【已修正】移除了单位和数值之间的空格。
     * @param actualValue 原始值
     * @param decimals    要保留的小数位数
     * @param unit        要附加的单位
     * @return 格式化后的字符串
     */
    public static String formatValue(String actualValue, Integer decimals, String unit) {
        if (actualValue == null || actualValue.trim().isEmpty()) {
            return "";
        }

        try {
            BigDecimal numberValue = new BigDecimal(actualValue);

            if (decimals != null && decimals >= 0) {
                numberValue = numberValue.setScale(decimals, RoundingMode.HALF_UP);
            }

            String formattedString = numberValue.toPlainString();

            if (unit != null && !unit.trim().isEmpty()) {
                // 【核心修改点 #1】移除此处的空格
                return formattedString + unit;
            } else {
                return formattedString;
            }

        } catch (NumberFormatException e) {
            // 如果无法解析为数字，则直接返回原始值，并附加单位（如果存在）
            if (unit != null && !unit.trim().isEmpty()) {
                // 【核心修改点 #2】移除此处的空格
                return actualValue + unit;
            }
            return actualValue;
        }
    }

    public static XSSFWorkbook createWorkbookFromTemplate(byte[] templateBytes) throws IOException {
        try (InputStream templateStream = new ByteArrayInputStream(templateBytes)) {
            return new XSSFWorkbook(templateStream);
        }
    }
}