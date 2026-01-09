package org.school.analysis.service.impl.report;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.config.AppConfig;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * БАЗОВЫЙ КЛАСС ДЛЯ ВСЕХ ГЕНЕРАТОРОВ ОТЧЕТОВ
 *
 * Содержит:
 * - Общие константы для настройки Excel
 * - Методы сохранения файлов
 * - Стили ячеек
 * - Утилиты форматирования
 */
@Slf4j
public abstract class ExcelReportBase {

    // ============ КОНФИГУРАЦИЯ ПУТЕЙ ============

    /** Путь для сохранения итоговых отчетов */
    protected static final String FINAL_REPORT_FOLDER = AppConfig.FINAL_REPORT_FOLDER;

    // ============ КОНФИГУРАЦИЯ КОЛОНОК ============

    /**
     * Минимальная ширина колонки в twips (1/20 точки)
     * Чем меньше значение, тем уже колонки
     * Рекомендуемый диапазон: 8-12 * 256
     */
    protected static final int MIN_COLUMN_WIDTH = 8 * 256;

    /**
     * Максимальная ширина колонки в twips
     * Чем больше значение, тем шире могут быть колонки
     * Рекомендуемый диапазон: 40-60 * 256
     */
    protected static final int MAX_COLUMN_WIDTH = 40 * 256;

    // ============ ШИРИНЫ ДЛЯ КОНКРЕТНЫХ ТИПОВ КОЛОНОК ============

    /** Ширина для колонки с номером студента */
    protected static final int COL_WIDTH_NUMBER = 1200;     // ~3 символа

    /** Ширина для колонки с ФИО студента */
    protected static final int COL_WIDTH_FIO = 4500;        // ~15-20 символов

    /** Ширина для колонки с присутствием */
    protected static final int COL_WIDTH_PRESENCE = 2000;   // ~6 символов

    /** Ширина для колонки с вариантом */
    protected static final int COL_WIDTH_VARIANT = 1500;    // ~4 символа

    /** Ширина для колонки с общим баллом */
    protected static final int COL_WIDTH_TOTAL_SCORE = 1500;

    /** Ширина для колонки с процентом выполнения */
    protected static final int COL_WIDTH_PERCENTAGE = 1500;

    /** Ширина для колонки с баллом за задание */
    protected static final int COL_WIDTH_TASK_SCORE = 800;  // ~2 символа

    // ============ НАСТРОЙКИ ЗАГОЛОВКОВ ============

    /** Размер шрифта заголовка отчета */
    protected static final short TITLE_FONT_SIZE = 16;

    /** Размер шрифта подзаголовка */
    protected static final short SUBTITLE_FONT_SIZE = 12;

    /** Размер шрифта заголовков секций */
    protected static final short SECTION_FONT_SIZE = 12;

    /** Размер шрифта заголовков таблиц */
    protected static final short TABLE_HEADER_FONT_SIZE = 11;

    /** Размер шрифта обычного текста */
    protected static final short NORMAL_FONT_SIZE = 10;

    // ============ ЦВЕТА ============

    /** Цвет фона заголовков таблиц (IndexedColors.LIGHT_CORNFLOWER_BLUE) */
    protected static final short TABLE_HEADER_BG_COLOR = IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex();

    /** Цвет фона заголовков секций (IndexedColors.GREY_25_PERCENT) */
    protected static final short SECTION_HEADER_BG_COLOR = IndexedColors.GREY_25_PERCENT.getIndex();

    /** Цвет фона итоговой строки (IndexedColors.DARK_BLUE) */
    protected static final short SUMMARY_ROW_BG_COLOR = IndexedColors.DARK_BLUE.getIndex();

    // ============ ОТСТУПЫ И ПРОБЕЛЫ ============

    /** Количество пустых строк между секциями */
    protected static final int SECTION_SPACING = 1;

    /** Количество пустых строк перед графиками */
    protected static final int CHART_SPACING = 2;

    // ============ МЕТОДЫ ДЛЯ РАБОТЫ С ФАЙЛАМИ ============

    /**
     * Создает папку для отчетов, если она не существует
     */
    protected Path createReportsFolder() throws IOException {
        Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
        if (!Files.exists(reportsPath)) {
            Files.createDirectories(reportsPath);
        }
        return reportsPath;
    }

    /**
     * Сохраняет рабочую книгу в файл
     */
    protected File saveWorkbook(Workbook workbook, Path folderPath, String fileName) throws IOException {
        Path filePath = folderPath.resolve(fileName);
        try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
            workbook.write(fos);
        }
        log.info("✅ Отчет сохранен: {}", filePath);
        return filePath.toFile();
    }

    // ============ МЕТОДЫ ДЛЯ ОПТИМИЗАЦИИ КОЛОНОК ============

    /**
     * Автоматически подбирает ширину колонок с ограничениями
     * @param sheet лист Excel
     * @param columnCount количество колонок для оптимизации
     */
    protected void optimizeColumnWidths(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
            int currentWidth = sheet.getColumnWidth(i);

            if (currentWidth < MIN_COLUMN_WIDTH) {
                sheet.setColumnWidth(i, MIN_COLUMN_WIDTH);
            } else if (currentWidth > MAX_COLUMN_WIDTH) {
                sheet.setColumnWidth(i, MAX_COLUMN_WIDTH);
            }
        }
    }

    /**
     * Устанавливает фиксированные ширины для колонок детального отчета
     * @param sheet лист Excel
     * @param maxTasks максимальное количество заданий
     */
    protected void optimizeDetailReportColumns(Sheet sheet, int maxTasks) {
        // Основные колонки (A-F)
        sheet.setColumnWidth(0, COL_WIDTH_NUMBER);      // №
        sheet.setColumnWidth(1, COL_WIDTH_FIO);         // ФИО
        sheet.setColumnWidth(2, COL_WIDTH_PRESENCE);    // Присутствие
        sheet.setColumnWidth(3, COL_WIDTH_VARIANT);     // Вариант
        sheet.setColumnWidth(4, COL_WIDTH_TOTAL_SCORE); // Общий балл
        sheet.setColumnWidth(5, COL_WIDTH_PERCENTAGE);  // % выполнения

        // Колонки с баллами за задания (G и далее)
        for (int i = 0; i < maxTasks; i++) {
            int colIndex = 6 + i; // Начинаем с колонки G (индекс 6)
            if (colIndex < 256) { // Максимум 256 колонок в Excel
                sheet.setColumnWidth(colIndex, COL_WIDTH_TASK_SCORE);
            }
        }
    }

    // ============ МЕТОДЫ ДЛЯ СОЗДАНИЯ СТИЛЕЙ ============

    /**
     * Создает стиль для заголовка отчета
     */
    protected CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(TITLE_FONT_SIZE);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для подзаголовка
     */
    protected CellStyle createSubtitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(SUBTITLE_FONT_SIZE);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для заголовков секций
     */
    protected CellStyle createSectionHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(SECTION_FONT_SIZE);
        style.setFont(font);
        style.setFillForegroundColor(SECTION_HEADER_BG_COLOR);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    /**
     * Создает стиль для заголовков таблиц
     */
    protected CellStyle createTableHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(TABLE_HEADER_FONT_SIZE);
        style.setFont(font);
        style.setFillForegroundColor(TABLE_HEADER_BG_COLOR);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для центрированных ячеек
     */
    protected CellStyle createCenteredStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для ячеек с процентами
     */
    protected CellStyle createPercentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для ячеек с десятичными числами
     */
    protected CellStyle createDecimalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для итоговых строк
     */
    protected CellStyle createSummaryStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(SUMMARY_ROW_BG_COLOR);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для обычного текста
     */
    protected CellStyle createNormalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints(NORMAL_FONT_SIZE);
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    // ============ УТИЛИТНЫЕ МЕТОДЫ ============

    /**
     * Устанавливает значение в ячейку с автоматическим определением типа
     */
    protected void setCellValue(Row row, int cellIndex, Object value) {
        Cell cell = row.createCell(cellIndex);

        if (value == null) {
            cell.setCellValue("");
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * Устанавливает числовое значение в ячейку
     */
    protected void setNumericCellValue(Row row, int cellIndex, Integer value) {
        Cell cell = row.createCell(cellIndex);
        if (value != null) {
            cell.setCellValue(value);
        } else {
            cell.setCellValue(0);
        }
    }

    /**
     * Добавляет пустые строки
     */
    protected int addEmptyRows(Sheet sheet, int currentRow, int count) {
        for (int i = 0; i < count; i++) {
            sheet.createRow(currentRow++);
        }
        return currentRow;
    }

    /**
     * Создает объединенную ячейку для заголовка
     */
    protected void createMergedTitle(Sheet sheet, String title, CellStyle style,
                                     int startRow, int startCol, int endCol) {
        Row row = sheet.createRow(startRow);
        Cell cell = row.createCell(startCol);
        cell.setCellValue(title);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, startCol, endCol));
    }
}