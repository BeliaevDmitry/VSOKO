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
     * Создает графики на основе существующей таблицы "АНАЛИЗ ПО ЗАДАНИЯМ"
     */
    public void createChartsFromAnalysisTable(XSSFWorkbook workbook, XSSFSheet sheet,
                                              int analysisTableStartRow, int taskCount,
                                              int chartStartRow) {

        try {
            log.debug("Создание графиков на основе таблицы анализа, задач: {}", taskCount);


            // Вычисляем позицию для графиков (после таблицы данных)
            int chartRow = chartStartRow + 2;

            int dataTableStartRow = analysisTableStartRow -1;
            // 1. Stacked Bar Chart
            stackedBarChartGenerator.createChartFromDataTable(
                    workbook, sheet,
                    dataTableStartRow, // строка с данными
                    taskCount,         // количество заданий
                    chartRow,          // позиция для графика
                    "Распределение результатов по заданиям"
            );

            chartRow += styleConfig.getRowSpan() + styleConfig.getSpacing();

// 2. Line Chart
            lineChartGenerator.createChartFromDataTable(
                    workbook, sheet,
                    dataTableStartRow,
                    taskCount,
                    chartRow,
                    "Процент выполнения заданий"
            );

            chartRow += styleConfig.getRowSpan() + styleConfig.getSpacing();

// 3. Percentage Bar Chart
            percentageBarChartGenerator.createChartFromDataTable(
                    workbook, sheet,
                    dataTableStartRow,
                    taskCount,
                    chartRow,
                    "Процент выполнения (столбчатая диаграмма)"
            );

            log.info("✅ Все графики успешно созданы на основе таблицы анализа");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании графиков из таблицы анализа: {}", e.getMessage(), e);
            Row errorRow = sheet.createRow(chartStartRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании графиков: " + e.getMessage());
        }
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