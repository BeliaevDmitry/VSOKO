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

    private final ChartStyleConfig styleConfig;
    private final StackedBarChartGenerator stackedBarChartGenerator;
    private final LineChartGenerator lineChartGenerator;
    private final PercentageBarChartGenerator percentageBarChartGenerator;

    /**
     * Создает все графики для детального отчета
     */
    public void createCharts(XSSFWorkbook workbook, XSSFSheet sheet,
                             TestSummaryDto testSummary,
                             Map<Integer, TaskStatisticsDto> taskStatistics,
                             int startRow) {

        try {
            log.debug("Создание графиков для теста: {}", testSummary.getFileName());

            // Создаем заголовок раздела с графиками
            Row sectionHeader = sheet.createRow(startRow++);
            Cell headerCell = sectionHeader.createCell(0);
            headerCell.setCellValue("ГРАФИЧЕСКИЙ АНАЛИЗ");
            headerCell.setCellStyle(createSectionHeaderStyle(workbook));

            startRow++; // Пустая строка

            if (taskStatistics == null || taskStatistics.isEmpty()) {
                log.warn("Нет данных для создания графиков");
                sheet.createRow(startRow++).createCell(0).setCellValue("Нет данных для графиков");
                return;
            }

            // Сортируем задания по номеру
            List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                    .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                    .collect(Collectors.toList());

            // Создаем таблицу данных для графиков
            int dataStartRow = createDataTable(workbook, sheet, sortedTasks, startRow);

            // Вычисляем позицию для графиков (после таблицы)
            int chartRow = dataStartRow + 2;

            // 1. Stacked Bar Chart - распределение результатов
            stackedBarChartGenerator.createChart(
                    workbook, sheet, sortedTasks,
                    dataStartRow, chartRow,
                    "Распределение результатов по заданиям"
            );

            chartRow += styleConfig.getRowSpan() + styleConfig.getSpacing();

            // 2. Line Chart - процент выполнения
            lineChartGenerator.createChart(
                    workbook, sheet, sortedTasks,
                    dataStartRow, chartRow,
                    "Процент выполнения заданий"
            );

            chartRow += styleConfig.getRowSpan() + styleConfig.getSpacing();

            // 3. Percentage Bar Chart - столбчатая диаграмма процентов
            percentageBarChartGenerator.createChart(
                    workbook, sheet, sortedTasks,
                    dataStartRow, chartRow,
                    "Процент выполнения (столбчатая диаграмма)"
            );

            log.info("✅ Все графики успешно созданы");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании графиков: {}", e.getMessage(), e);
            Row errorRow = sheet.createRow(startRow++);
            errorRow.createCell(0).setCellValue("Ошибка при создании графиков: " + e.getMessage());
        }
    }

    /**
     * Создает таблицу данных для графиков
     */
    private int createDataTable(Workbook workbook, Sheet sheet,
                                List<TaskStatisticsDto> tasks, int startRow) {

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

        return startRow - tasks.size() - 1; // Возвращаем начальную строку данных
    }

    /**
     * Создает стиль для заголовка секции
     */
    private CellStyle createSectionHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
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