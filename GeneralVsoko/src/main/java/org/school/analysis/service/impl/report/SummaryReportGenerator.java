package org.school.analysis.service.impl.report;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

@Component
@Slf4j
public class SummaryReportGenerator extends ExcelReportBase {

    public File generateSummaryReport(List<TestSummaryDto> tests, String schoolName) {
        log.info("Генерация сводного отчета для {} тестов", tests.size());

        try {
            Path reportsPath = createReportsFolder(schoolName);

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Все тесты");

                createHeader(sheet, workbook, tests);
                fillData(sheet, workbook, tests);
                optimizeColumnWidths(sheet, 18);

                return saveWorkbook(workbook, reportsPath, "Свод всех работ.xlsx");
            }

        } catch (IOException e) {
            log.error("Ошибка при создании сводного отчета", e);
            throw new RuntimeException("Ошибка при создании отчета", e);
        }
    }

    private void createHeader(Sheet sheet, Workbook workbook, List<TestSummaryDto> tests) {
        // Используем кэшированные стили
        CellStyle titleStyle = getTitleStyle(workbook);
        CellStyle tableHeaderStyle = getTableHeaderStyle(workbook);
        CellStyle infoStyle = getSubtitleStyle(workbook);

        // 1. Заголовок отчета
        createMergedTitle(sheet,
                "СВОДНЫЙ ОТЧЕТ ПО ВСЕМ ТЕСТАМ",
                titleStyle,
                0, 0, HEADER_MERGE_COUNT_SUMMARY_REPORT);

        // 2. Строка с информацией
        Row infoRow = sheet.createRow(1);

        // Вариант 1: Обе надписи в одной строке
        // Дата формирования (колонки A-B)
        Cell dateCell = infoRow.createCell(0);
        dateCell.setCellValue("Дата формирования: " +
                LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE));
        dateCell.setCellStyle(infoStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 2)); // A-C

        // Количество тестов (колонки D-E)
        Cell summaryCell = infoRow.createCell(3); // Колонка D
        summaryCell.setCellValue("Тестов: " + tests.size());
        summaryCell.setCellStyle(infoStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 5)); // D-F

        // 3. Пустая строка для визуального разделения
        sheet.createRow(2);

        // 4. Заголовки колонок таблицы
        String[] headers = {
                "Учебный год", "Предмет", "Класс", "Учитель", "Дата теста", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения", "Кол-во заданий", "Макс. балл"
        };

        Row headerRow = sheet.createRow(3);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(tableHeaderStyle);
        }

        // 5. Включаем фильтры
        setupTableFeatures(sheet, headerRow, headers.length);
    }

    /**
     * Настройка всех функций таблицы
     */
    private void setupTableFeatures(Sheet sheet, Row headerRow, int columnCount) {
        // 1. Автофильтр
        enableAutoFilter(sheet, headerRow, columnCount);

        // 2. Закрепление областей - 4 строки (0-3)
        createFreezePane(sheet, 4);

        // 3. Автоподбор ширины колонок для заголовков
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }

        // 4. Создаем именованный диапазон для таблицы (опционально)
        createNamedRange(sheet, headerRow, columnCount);
    }

    /**
     * Закрепляет строки (создает freeze pane)
     */
    private void createFreezePane(Sheet sheet, int rowsToFreeze) {
        try {
            // Закрепляем первые N строк и 0 колонок
            // Параметры: колонки для закрепления слева, строки для закрепления сверху
            sheet.createFreezePane(0, rowsToFreeze); // 0 колонок, N строк сверху

            log.debug("Закреплены первые {} строк", rowsToFreeze);

            // Проверка (опционально)
            if (sheet instanceof org.apache.poi.xssf.usermodel.XSSFSheet) {
                org.apache.poi.xssf.usermodel.XSSFSheet xssfSheet = (org.apache.poi.xssf.usermodel.XSSFSheet) sheet;
                log.debug("Freeze pane установлен на строку: {}", xssfSheet.getPaneInformation().getVerticalSplitPosition());
            }

        } catch (Exception e) {
            log.error("Ошибка при закреплении строк: {}", e.getMessage());
            throw new RuntimeException("Не удалось закрепить строки в Excel", e);
        }
    }


    /**
     * Создание именованного диапазона для удобства
     */
    private void createNamedRange(Sheet sheet, Row headerRow, int columnCount) {
        Workbook workbook = sheet.getWorkbook();

        // Имя для диапазона данных
        String rangeName = "TableData";

        // Формула диапазона: от заголовков до последней строки
        String formula = String.format("'%s'!$A$%d:$%s$%d",
                sheet.getSheetName(),
                headerRow.getRowNum() + 1, // +1 потому что Excel формулы 1-based
                getExcelColumnName(columnCount - 1),
                sheet.getLastRowNum() + 1);

        try {
            Name namedRange = workbook.createName();
            namedRange.setNameName(rangeName);
            namedRange.setRefersToFormula(formula);
            log.debug("Создан именованный диапазон '{}': {}", rangeName, formula);
        } catch (Exception e) {
            log.warn("Не удалось создать именованный диапазон: {}", e.getMessage());
        }
    }
    /**
     * Включает автофильтр на заголовках таблицы
     */
    private void enableAutoFilter(Sheet sheet, Row headerRow, int columnCount) {
        // Устанавливаем диапазон для автофильтра
        // Фильтр будет от строки 3 (заголовки) до последней строки с данными
        CellRangeAddress filterRange = new CellRangeAddress(
                headerRow.getRowNum(), // первая строка (заголовки)
                sheet.getLastRowNum(), // последняя строка (данные)
                0,                    // первая колонка
                columnCount - 1       // последняя колонка
        );

        sheet.setAutoFilter(filterRange);
        log.debug("Автофильтр установлен на диапазоне: строка {}-{}, колонки {}-{}",
                filterRange.getFirstRow(), filterRange.getLastRow(),
                filterRange.getFirstColumn(), filterRange.getLastColumn());
    }


    /**
     * Преобразует индекс колонки в буквенное обозначение Excel (A, B, C, ... AA, AB, ...)
     */
    private String getExcelColumnName(int columnIndex) {
        StringBuilder columnName = new StringBuilder();

        while (columnIndex >= 0) {
            int remainder = columnIndex % 26;
            columnName.insert(0, (char) ('A' + remainder));
            columnIndex = (columnIndex / 26) - 1;
        }

        return columnName.toString();
    }

    private void fillData(Sheet sheet, Workbook workbook, List<TestSummaryDto> tests) {
        // 1. СОЗДАЕМ СТИЛИ ОДИН РАЗ (из кэша)
        CellStyle normalStyle = getStyle(workbook, StyleType.NORMAL);
        CellStyle percentStyle = getStyle(workbook, StyleType.PERCENT);
        CellStyle decimalStyle = getStyle(workbook, StyleType.DECIMAL);
        CellStyle centeredStyle = getStyle(workbook, StyleType.CENTERED);

        // 2. Определяем типы данных для колонок
        CellStyle[] columnStyles = {
                normalStyle,    // 0: Учебный год (текст слева)
                normalStyle,    // 1: Предмет
                normalStyle,    // 2: Класс
                normalStyle,    // 3: Учитель
                centeredStyle,  // 4: Дата теста
                centeredStyle,  // 5: Тип теста
                centeredStyle,  // 6: Присутствовало
                centeredStyle,  // 7: Отсутствовало
                centeredStyle,  // 8: Всего
                percentStyle,   // 9: % присутствия
                decimalStyle,   // 10: Средний балл
                percentStyle,   // 11: % выполнения
                centeredStyle,  // 12: Кол-во заданий
                centeredStyle   // 13: Макс. балл
        };

        int rowNum = 4;

        for (TestSummaryDto test : tests) {
            Row row = sheet.createRow(rowNum++);

            // 3. Заполняем данные с правильными стилями
            // Текстовые поля
            setCellValue(row, 0, test.getAcademicYear(), columnStyles[0]);
            setCellValue(row, 1, test.getSubject(), columnStyles[1]);
            setCellValue(row, 2, test.getClassName(), columnStyles[2]);
            setCellValue(row, 3, test.getTeacher(), columnStyles[3]);

            // Дата
            setCellValue(row, 4,
                    test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE),
                    columnStyles[4]);

            // Тип теста
            setCellValue(row, 5, test.getTestType(), columnStyles[5]);

            // Числовые значения
            setCellValue(row, 6, test.getStudentsPresent(), columnStyles[6]);
            setCellValue(row, 7, test.getStudentsAbsent(), columnStyles[7]);
            setCellValue(row, 8, test.getClassSize(), columnStyles[8]);

            // Проценты (конвертируем в double)
            if (test.getAttendancePercentage() != null) {
                setCellValue(row, 9, test.getAttendancePercentage() / 100.0, columnStyles[9]);
            } else {
                setCellValue(row, 9, 0.0, columnStyles[9]);
            }

            // Десятичные числа
            setCellValue(row, 10, test.getAverageScore() != null ?
                    test.getAverageScore() : 0.0, columnStyles[10]);

            // Проценты выполнения
            if (test.getSuccessPercentage() != null) {
                setCellValue(row, 11, test.getSuccessPercentage() / 100.0, columnStyles[11]);
            } else {
                setCellValue(row, 11, 0.0, columnStyles[11]);
            }

            // Остальные числовые значения
            setCellValue(row, 12, test.getTaskCount(), columnStyles[12]);
            setCellValue(row, 13, test.getMaxTotalScore(), columnStyles[13]);
        }
    }
}