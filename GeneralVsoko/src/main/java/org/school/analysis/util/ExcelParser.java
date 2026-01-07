package org.school.analysis.util;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;

/**
 * Утилитарный класс для низкоуровневого парсинга Excel ячеек
 */
public class ExcelParser {

    private static final DateTimeFormatter[] DATE_FORMATS = {
            DateTimeFormatter.ofPattern("dd.MM.yyyy"),
            DateTimeFormatter.ofPattern("dd/MM/yyyy"),
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("dd.MM.yy"),
            DateTimeFormatter.ofPattern("d.M.yyyy")
    };

    /**
     * Получение строкового значения из ячейки
     */
    public static String getCellValueAsString(Cell cell) {
        return getCellValueAsString(cell, null);
    }

    public static String getCellValueAsString(Cell cell, String defaultValue) {
        if (cell == null) {
            return defaultValue;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Обработка даты
                    try {
                        return cell.getDateCellValue().toInstant()
                                .atZone(java.time.ZoneId.systemDefault())
                                .toLocalDate()
                                .format(DateTimeFormatter.ISO_LOCAL_DATE);
                    } catch (Exception e) {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                } else {
                    // Обработка чисел
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value) && !Double.isInfinite(value)) {
                        // Целое число
                        return String.valueOf((int) value);
                    } else {
                        // Дробное число
                        String str = String.valueOf(value);
                        // Убираем ненужные .0
                        if (str.endsWith(".0")) {
                            return str.substring(0, str.length() - 2);
                        }
                        return str;
                    }
                }

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return defaultValue;
                    }
                }

            case BLANK:
                return defaultValue;

            default:
                return defaultValue;
        }
    }

    /**
     * Получение значения из ячейки по координатам листа
     */
    public static String getCellValueAsString(Sheet sheet, int rowIndex, int colIndex,
                                              String defaultValue) {
        if (sheet == null) {
            return defaultValue;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return defaultValue;
        }

        return getCellValueAsString(row.getCell(colIndex), defaultValue);
    }

    /**
     * Перегруженный вариант без defaultValue (возвращает null)
     */
    public static String getCellValueAsString(Sheet sheet, int rowIndex, int colIndex) {
        return getCellValueAsString(sheet, rowIndex, colIndex, null);
    }

    /**
     * Получение целочисленного значения из ячейки
     */
    public static Integer getCellValueAsInteger(Cell cell) {
        return getCellValueAsInteger(cell, null);
    }

    public static Integer getCellValueAsInteger(Cell cell, Integer defaultValue) {
        String stringValue = getCellValueAsString(cell);
        if (stringValue == null || stringValue.trim().isEmpty()) {
            return defaultValue;
        }

        try {
            // Убираем возможную дробную часть
            stringValue = stringValue.replaceAll("\\.0+$", "");
            // Заменяем запятые на точки для десятичных чисел
            stringValue = stringValue.replace(",", ".");

            if (stringValue.contains(".")) {
                // Если десятичное, округляем
                return (int) Math.round(Double.parseDouble(stringValue));
            } else {
                // Целое число
                return Integer.parseInt(stringValue);
            }
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    /**
     * Получение значения с плавающей точкой
     */
    public static Double getCellValueAsDouble(Cell cell) {
        return getCellValueAsDouble(cell, null);
    }

    public static Double getCellValueAsDouble(Cell cell, Double defaultValue) {
        String stringValue = getCellValueAsString(cell);
        if (stringValue == null || stringValue.trim().isEmpty()) {
            return defaultValue;
        }

        try {
            stringValue = stringValue.replace(",", ".");
            return Double.parseDouble(stringValue);
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    /**
     * Получение даты из ячейки
     */
    public static LocalDate getCellValueAsDate(Cell cell) {
        return getCellValueAsDate(cell, null);
    }

    public static LocalDate getCellValueAsDate(Cell cell, LocalDate defaultValue) {
        if (cell == null) {
            return defaultValue;
        }

        // Если ячейка отформатирована как дата
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            try {
                return cell.getDateCellValue().toInstant()
                        .atZone(java.time.ZoneId.systemDefault())
                        .toLocalDate();
            } catch (Exception e) {
                // Пробуем парсить как строку
            }
        }

        // Пробуем парсить как строку
        String stringValue = getCellValueAsString(cell);
        if (stringValue == null) {
            return defaultValue;
        }

        for (DateTimeFormatter formatter : DATE_FORMATS) {
            try {
                return LocalDate.parse(stringValue, formatter);
            } catch (DateTimeParseException e) {
                // Пробуем следующий формат
            }
        }

        return defaultValue;
    }

    /**
     * Получение логического значения
     */
    public static Boolean getCellValueAsBoolean(Cell cell) {
        return getCellValueAsBoolean(cell, null);
    }

    public static Boolean getCellValueAsBoolean(Cell cell, Boolean defaultValue) {
        if (cell == null) {
            return defaultValue;
        }

        if (cell.getCellType() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        }

        String stringValue = getCellValueAsString(cell);
        if (stringValue == null) {
            return defaultValue;
        }

        stringValue = stringValue.trim().toLowerCase();
        return stringValue.equals("true") ||
                stringValue.equals("да") ||
                stringValue.equals("yes") ||
                stringValue.equals("1");
    }

    /**
     * Проверка, пуста ли ячейка
     */
    public static boolean isCellEmpty(Cell cell) {
        if (cell == null) {
            return true;
        }

        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim().isEmpty();
        }

        return false;
    }
}