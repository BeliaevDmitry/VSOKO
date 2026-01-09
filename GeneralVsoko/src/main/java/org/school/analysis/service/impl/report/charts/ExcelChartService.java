package org.school.analysis.service.impl.report.charts;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.springframework.stereotype.Service;

import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class ExcelChartService {

    private final StackedBarChartGenerator stackedBarChartGenerator;
    private final PercentageBarChartGenerator percentageBarChartGenerator;
    private final LineChartGenerator lineChartGenerator;

    /**
     * Создает все графики для детального отчета
     */
    public void createCharts(XSSFWorkbook workbook, Sheet sheet,
                             TestSummaryDto testSummary,
                             Map<Integer, TaskStatisticsDto> taskStatistics,
                             int startRow) {

        try {
            log.debug("Создание графиков для теста: {}", testSummary.getFileName());

            // Создаем заголовок раздела с графиками
            Row sectionHeader = sheet.createRow(startRow++);
            sectionHeader.createCell(0).setCellValue("ГРАФИЧЕСКИЙ АНАЛИЗ");
            sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            startRow++; // Пустая строка

            if (taskStatistics == null || taskStatistics.isEmpty()) {
                log.warn("Нет данных для создания графиков");
                return;
            }

            // Сортируем задания по номеру
            List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                    .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                    .collect(Collectors.toList());

            // Создаем таблицу данных для графиков
            int dataStartRow = createChartDataTable(workbook, sheet, sortedTasks, startRow);

            int chartStartRow = dataStartRow + 2; // Отступ после таблицы данных

            // Создаем графики последовательно
            createChartSection(workbook, sheet, sortedTasks, dataStartRow, chartStartRow);

            log.info("✅ Все графики успешно созданы");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании графиков: {}", e.getMessage(), e);
        }
    }

    /**
     * Создает раздел со всеми графиками
     */
    private void createChartSection(XSSFWorkbook workbook, Sheet sheet,
                                    List<TaskStatisticsDto> sortedTasks,
                                    int dataStartRow, int chartStartRow) {

        // 1. Stacked Bar Chart
        stackedBarChartGenerator.createChart(
                workbook, (XSSFSheet) sheet, sortedTasks,
                dataStartRow, chartStartRow,
                "Распределение результатов (Stacked)"
        );

        // 2. Percentage Bar Chart
        chartStartRow += 20; // Отступ между графиками
        percentageBarChartGenerator.createChart(
                workbook, (XSSFSheet) sheet, sortedTasks,
                dataStartRow, chartStartRow,
                "Процент выполнения заданий"
        );

        // 3. Line Chart
        chartStartRow += 20; // Отступ между графиками
        lineChartGenerator.createChart(
                workbook, (XSSFSheet) sheet, sortedTasks,
                dataStartRow, chartStartRow,
                "Динамика выполнения заданий"
        );
    }

    /**
     * Создает таблицу данных для графиков
     */
    private int createChartDataTable(Workbook workbook, Sheet sheet,
                                     List<TaskStatisticsDto> tasks, int startRow) {

        // Проверяем, не перекрывает ли таблица существующие данные
        int lastRowWithData = sheet.getLastRowNum();
        if (startRow <= lastRowWithData) {
            startRow = lastRowWithData + 2;
        }

        // Заголовки таблицы
        Row headerRow = sheet.createRow(startRow++);
        String[] headers = {"Задание", "Полностью", "Частично", "Не справилось", "% выполнения"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        // Данные для каждого задания
        for (TaskStatisticsDto task : tasks) {
            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue("№" + task.getTaskNumber());
            row.createCell(1).setCellValue(task.getFullyCompletedCount());
            row.createCell(2).setCellValue(task.getPartiallyCompletedCount());
            row.createCell(3).setCellValue(task.getNotCompletedCount());

            Cell percentCell = row.createCell(4);
            percentCell.setCellValue(task.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(createPercentStyle(workbook));
        }

        return startRow; // Возвращаем следующую свободную строку
    }

    /**
     * Создает стиль для заголовка секции
     */
    private CellStyle createSectionHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
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
    private CellStyle createTableHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для ячеек с процентами
     */
    private CellStyle createPercentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }
}