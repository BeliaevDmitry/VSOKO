package org.school.analysis.service.impl;

import org.school.analysis.model.dto.TeacherTestDetailDto;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.ExcelReportService;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class ExcelReportServiceImpl implements ExcelReportService {

    private static final String FINAL_REPORT_FOLDER = AppConfig.FINAL_REPORT_FOLDER;
    private static final int CHART_COL_SPAN = 8;
    private static final int CHART_ROW_SPAN = 20;
    private static final int MIN_COLUMN_WIDTH = 10 * 256;
    private static final int MAX_COLUMN_WIDTH = 50 * 256;

    @Override
    public File generateSummaryReport(List<TestSummaryDto> tests) {
        log.info("Генерация сводного отчета для {} тестов", tests.size());

        try {
            Path reportsPath = createReportsFolder();

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                XSSFSheet sheet = workbook.createSheet("Сводка по тестам");

                createSummarySheetHeader(sheet, workbook);
                fillSummaryData(sheet, workbook, tests);
                addSummaryStatisticsRow(sheet, tests);
                optimizeColumnWidths(sheet);

                return saveWorkbook(workbook, reportsPath, "Свод всех работ.xlsx");
            }

        } catch (IOException e) {
            log.error("Ошибка при создании сводного отчета", e);
            throw new RuntimeException("Ошибка при создании отчета", e);
        }
    }

    @Override
    public File generateTestDetailReport(
            TestSummaryDto testSummary,
            List<StudentDetailedResultDto> studentResults,
            Map<Integer, TaskStatisticsDto> taskStatistics) {

        log.info("Генерация детального отчета на одном листе для теста: {} - {}",
                testSummary.getSubject(), testSummary.getClassName());

        try {
            Path reportsPath = createReportsFolder();

            String fileName = String.format("Детальный_отчет_%s_%s_%s.xlsx",
                    testSummary.getSubject(),
                    testSummary.getClassName(),
                    testSummary.getTestDate().format(DateTimeFormatter.ofPattern("ddMMyyyy")));

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                String sheetName = String.format("%s_%s",
                        testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", ""),
                        testSummary.getClassName());

                createSingleSheetDetailReport(workbook, testSummary, studentResults, taskStatistics, sheetName);

                return saveWorkbook(workbook, reportsPath, fileName);
            }

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета", e);
            return null;
        }
    }

    @Override
    public File generateTeacherReport(
            String teacherName,
            List<TestSummaryDto> teacherTests) {

        log.info("Генерация отчета для учителя: {}", teacherName);

        try {
            Path reportsPath = createReportsFolder();

            String fileName = String.format("Отчет_учителя_%s_%s.xlsx",
                    teacherName.replace(" ", "_"),
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmm")));

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                createTeacherSummarySheet(workbook, teacherName, teacherTests);

                for (int i = 0; i < teacherTests.size(); i++) {
                    createTeacherTestSheet(workbook, teacherTests.get(i), teacherName, i);
                }

                return saveWorkbook(workbook, reportsPath, fileName);
            }

        } catch (Exception e) {
            log.error("Ошибка при создании отчета учителя", e);
            return null;
        }
    }

    @Override
    public File generateTeacherReportWithDetails(String teacherName,
                                                 List<TestSummaryDto> teacherTests,
                                                 List<TeacherTestDetailDto> teacherTestDetails) {

        log.info("Генерация детального отчета для учителя: {} ({} тестов)",
                teacherName, teacherTests.size());

        try {
            Path reportsPath = createReportsFolder();

            String fileName = String.format("Отчет_учителя_%s_%s.xlsx",
                    teacherName.replace(" ", "_"),
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmm")));

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                createTeacherSummarySheet(workbook, teacherName, teacherTests);

                for (int i = 0; i < teacherTestDetails.size(); i++) {
                    TeacherTestDetailDto testDetail = teacherTestDetails.get(i);
                    createCompleteTestDetailSheetForTeacher(workbook, testDetail, teacherName, i);
                }

                return saveWorkbook(workbook, reportsPath, fileName);
            }

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета учителя", e);
            return null;
        }
    }

    private void createCompleteTestDetailSheetForTeacher(XSSFWorkbook workbook,
                                                         TeacherTestDetailDto testDetail,
                                                         String teacherName,
                                                         int sheetIndex) {

        TestSummaryDto testSummary = testDetail.getTestSummary();
        if (testSummary == null) {
            log.warn("Пропускаем тест без основных данных");
            return;
        }

        String sheetName = String.format("%s_%s_%s",
                testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", "").substring(0,
                        Math.min(12, testSummary.getSubject().length())),
                testSummary.getClassName(),
                testSummary.getTestDate().format(DateTimeFormatter.ofPattern("ddMM")));

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        String finalSheetName = sheetName;
        int counter = 1;
        while (workbook.getSheet(finalSheetName) != null) {
            finalSheetName = String.format("%s_%d", sheetName, counter++);
            if (finalSheetName.length() > 31) {
                finalSheetName = finalSheetName.substring(0, 31);
            }
        }

        try {
            createSingleSheetDetailReport(workbook, testSummary,
                    testDetail.getStudentResults(),
                    testDetail.getTaskStatistics(),
                    finalSheetName);

        } catch (Exception e) {
            log.error("Ошибка при создании детального листа для теста {}: {}",
                    testSummary.getFileName(), e.getMessage(), e);
        }
    }

    private void createSingleSheetDetailReport(XSSFWorkbook workbook,
                                               TestSummaryDto testSummary,
                                               List<StudentDetailedResultDto> studentResults,
                                               Map<Integer, TaskStatisticsDto> taskStatistics,
                                               String sheetName) {

        XSSFSheet sheet = workbook.createSheet(sheetName);

        int currentRow = 0;

        try {
            currentRow = createReportHeader(sheet, workbook, testSummary, currentRow);
            currentRow = createTestStatisticsSection(sheet, workbook, testSummary, currentRow);
            currentRow = createStudentResultsSection(sheet, workbook, studentResults, currentRow);
            currentRow = createTaskAnalysisSection(sheet, workbook, taskStatistics, currentRow);

            if (taskStatistics != null && !taskStatistics.isEmpty() &&
                    studentResults != null && !studentResults.isEmpty()) {
                createExcelChartsSection(workbook, sheet, testSummary, taskStatistics, studentResults, currentRow);
            }

            optimizeColumnWidths(sheet);

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета на одном листе", e);
            XSSFRow errorRow = sheet.createRow(currentRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании отчета: " + e.getMessage());
        }
    }

    private int createReportHeader(Sheet sheet, Workbook workbook,
                                   TestSummaryDto testSummary, int startRow) {
        Row titleRow = sheet.createRow(startRow++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(
                String.format("ДЕТАЛЬНЫЙ ОТЧЕТ ПО ТЕСТУ: %s - %s",
                        testSummary.getSubject(), testSummary.getClassName())
        );
        CellStyle titleStyle = createTitleStyle(workbook);
        titleCell.setCellStyle(titleStyle);

        Row subtitleRow = sheet.createRow(startRow++);
        Cell subtitleCell = subtitleRow.createCell(0);
        subtitleCell.setCellValue(
                String.format("Дата проведения: %s | Тип работы: %s",
                        testSummary.getTestDate().format(DateTimeFormatters.DISPLAY_DATE),
                        testSummary.getTestType())
        );
        CellStyle subtitleStyle = createSubtitleStyle(workbook);
        subtitleCell.setCellStyle(subtitleStyle);

        sheet.createRow(startRow++).createCell(0).setCellValue("Учитель: " + testSummary.getTeacher());
        sheet.createRow(startRow++).createCell(0).setCellValue("Файл отчета: " + testSummary.getFileName());

        startRow++;
        return startRow;
    }

    private int createTestStatisticsSection(Sheet sheet, Workbook workbook,
                                            TestSummaryDto testSummary, int startRow) {
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("ОБЩАЯ СТАТИСТИКА ТЕСТА");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        String[][] statsData = {
                {"Всего учеников в классе", String.valueOf(testSummary.getClassSize())},
                {"Присутствовало на тесте", String.valueOf(testSummary.getStudentsPresent())},
                {"Отсутствовало на тесте", String.valueOf(testSummary.getStudentsAbsent())},
                {"Процент присутствия", String.format("%.1f%%", testSummary.getAttendancePercentage())},
                {"", ""},
                {"Количество заданий", String.valueOf(testSummary.getTaskCount())},
                {"Максимальный балл за тест", String.valueOf(testSummary.getMaxTotalScore())},
                {"Средний балл по классу", String.format("%.2f", testSummary.getAverageScore())},
                {"Процент выполнения теста", String.format("%.1f%%", testSummary.getSuccessPercentage())}
        };

        for (String[] rowData : statsData) {
            Row row = sheet.createRow(startRow++);
            row.createCell(0).setCellValue(rowData[0]);
            row.createCell(1).setCellValue(rowData[1]);
        }

        startRow++;
        return startRow;
    }

    private int createStudentResultsSection(Sheet sheet, Workbook workbook,
                                            List<StudentDetailedResultDto> studentResults,
                                            int startRow) {
        if (studentResults == null || studentResults.isEmpty()) {
            return startRow;
        }

        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("РЕЗУЛЬТАТЫ СТУДЕНТОВ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        int maxTasks = studentResults.stream()
                .map(StudentDetailedResultDto::getTaskScores)
                .filter(Objects::nonNull)
                .mapToInt(Map::size)
                .max()
                .orElse(0);

        List<String> headers = new ArrayList<>(Arrays.asList(
                "№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"
        ));

        for (int i = 1; i <= maxTasks; i++) {
            headers.add("№" + i); // Изменено здесь с "З" + i на "№" + i
        }

        Row headerRow = sheet.createRow(startRow++);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        for (int i = 0; i < studentResults.size(); i++) {
            StudentDetailedResultDto student = studentResults.get(i);
            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue(i + 1);
            setCellValue(row, 1, student.getFio());
            setCellValue(row, 2, student.getPresence());
            setCellValue(row, 3, student.getVariant());
            setCellValue(row, 4, student.getTotalScore());

            Cell percentCell = row.createCell(5);
            double percentage = student.getPercentageScore() != null ?
                    student.getPercentageScore() / 100.0 : 0.0;
            percentCell.setCellValue(percentage);
            percentCell.setCellStyle(createPercentStyle(workbook));

            Map<Integer, Integer> taskScores = student.getTaskScores();
            for (int taskNum = 1; taskNum <= maxTasks; taskNum++) {
                Integer score = taskScores != null ? taskScores.get(taskNum) : null;
                Cell scoreCell = row.createCell(5 + taskNum);
                if (score != null) {
                    scoreCell.setCellValue(score);
                } else {
                    scoreCell.setCellValue(0);
                }
                scoreCell.setCellStyle(createCenteredStyle(workbook));
            }
        }

        startRow++;
        return startRow;
    }

    private int createTaskAnalysisSection(Sheet sheet, Workbook workbook,
                                          Map<Integer, TaskStatisticsDto> taskStatistics,
                                          int startRow) {
        if (taskStatistics == null || taskStatistics.isEmpty()) {
            return startRow;
        }

        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("АНАЛИЗ ПО ЗАДАНИЯМ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        String[] headers = {
                "№ задания", "Макс. балл", "Полностью", "Частично", "Не справилось",
                "% выполнения", "Распределение баллов"
        };

        Row headerRow = sheet.createRow(startRow++);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        for (TaskStatisticsDto stats : taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList())) {
            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue("№" + stats.getTaskNumber()); // Изменено здесь
            row.createCell(1).setCellValue(stats.getMaxScore());
            row.createCell(2).setCellValue(stats.getFullyCompletedCount());
            row.createCell(3).setCellValue(stats.getPartiallyCompletedCount());
            row.createCell(4).setCellValue(stats.getNotCompletedCount());

            Cell percentCell = row.createCell(5);
            percentCell.setCellValue(stats.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(createPercentStyle(workbook));

            if (stats.getScoreDistribution() != null) {
                String distribution = stats.getScoreDistribution().entrySet().stream()
                        .sorted(Map.Entry.comparingByKey())
                        .map(e -> String.format("%d баллов: %d", e.getKey(), e.getValue()))
                        .collect(Collectors.joining("; "));
                row.createCell(6).setCellValue(distribution);
            }
        }

        startRow++;
        return startRow;
    }

    /**
     * Создает нормированную гистограмму с накоплением - ОДИН столбец на задание
     */
    private void createSingleColumnNormalizedChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                                   List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            // Сначала создаем таблицу с нормированными данными для каждого задания
            int normalizedDataStartRow = chartRow + 2; // Отступаем немного от заголовка

            // Создаем заголовок таблицы с нормированными данными
            Row normHeaderRow = sheet.createRow(normalizedDataStartRow++);
            normHeaderRow.createCell(0).setCellValue("Задание");
            normHeaderRow.createCell(1).setCellValue("Полностью (%)");
            normHeaderRow.createCell(2).setCellValue("Частично (%)");
            normHeaderRow.createCell(3).setCellValue("Не справилось (%)");

            // Заполняем нормированные данные
            for (int i = 0; i < tasks.size(); i++) {
                TaskStatisticsDto task = tasks.get(i);
                Row row = sheet.createRow(normalizedDataStartRow++);

                row.createCell(0).setCellValue("№" + task.getTaskNumber());

                // Вычисляем общее количество студентов для этого задания
                int totalStudents = task.getFullyCompletedCount() +
                        task.getPartiallyCompletedCount() +
                        task.getNotCompletedCount();

                if (totalStudents > 0) {
                    // Нормируем данные в проценты
                    Cell fullyCell = row.createCell(1);
                    double fullyPercent = (double) task.getFullyCompletedCount() / totalStudents;
                    fullyCell.setCellValue(fullyPercent);
                    fullyCell.setCellStyle(createPercentStyle(workbook));

                    Cell partiallyCell = row.createCell(2);
                    double partiallyPercent = (double) task.getPartiallyCompletedCount() / totalStudents;
                    partiallyCell.setCellValue(partiallyPercent);
                    partiallyCell.setCellStyle(createPercentStyle(workbook));

                    Cell notCompletedCell = row.createCell(3);
                    double notCompletedPercent = (double) task.getNotCompletedCount() / totalStudents;
                    notCompletedCell.setCellValue(notCompletedPercent);
                    notCompletedCell.setCellStyle(createPercentStyle(workbook));
                } else {
                    row.createCell(1).setCellValue(0);
                    row.createCell(2).setCellValue(0);
                    row.createCell(3).setCellValue(0);
                }
            }

            // Теперь создаем диаграмму
            Row chartTitleRow = sheet.createRow(chartRow);
            chartTitleRow.createCell(0).setCellValue("Нормированное распределение результатов по заданиям (%)");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            // Увеличиваем отступ для диаграммы, чтобы она была после таблицы
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    0, normalizedDataStartRow + 2, CHART_COL_SPAN, normalizedDataStartRow + 2 + CHART_ROW_SPAN);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Нормированное распределение результатов (%)");

            // Получаем легенду
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("№ задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Доля студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            leftAxis.setMinimum(0.0);
            leftAxis.setMaximum(1.0); // 100%

            // Ключевое изменение: используем STACKED (не PERCENT_STACKED)
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STACKED); // Изменено с PERCENT_STACKED на STACKED
            barData.setVaryColors(true);

            // Используем нормированные данные из созданной таблицы
            CellRangeAddress labelRange = new CellRangeAddress(
                    chartRow + 3, // Начинаем с первой строки данных (после заголовка)
                    chartRow + 2 + tasks.size(), // Заканчиваем последней строкой данных
                    0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            // Используем нормированные данные (проценты)
            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 1, 1));
            XDDFChartData.Series series1 = data.addSeries(xs, ys1);
            series1.setTitle("Полностью", null);

            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 2, 2));
            XDDFChartData.Series series2 = data.addSeries(xs, ys2);
            series2.setTitle("Частично", null);

            XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 3, 3));
            XDDFChartData.Series series3 = data.addSeries(xs, ys3);
            series3.setTitle("Не справилось", null);

            chart.plot(data);

            // Настраиваем цвета
            setSeriesColor(chart, 0, new Color(46, 204, 113)); // Зеленый
            setSeriesColor(chart, 1, new Color(241, 196, 15)); // Желтый
            setSeriesColor(chart, 2, new Color(231, 76, 60));  // Красный

            // Настраиваем ось
            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

            log.info("✅ Нормированная гистограмма с накоплением (один столбец) создана успешно");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании нормированной гистограммы: {}", e.getMessage(), e);
        }
    }

    private void createExcelChartsSection(XSSFWorkbook workbook, Sheet sheet,
                                          TestSummaryDto testSummary,
                                          Map<Integer, TaskStatisticsDto> taskStatistics,
                                          List<StudentDetailedResultDto> studentResults,
                                          int startRow) {
        try {
            Row sectionHeader = sheet.createRow(startRow++);
            sectionHeader.createCell(0).setCellValue("ГРАФИЧЕСКИЙ АНАЛИЗ");
            sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            if (taskStatistics == null || taskStatistics.isEmpty()) {
                return;
            }

            List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                    .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                    .collect(Collectors.toList());

            int dataStartRow = startRow;
            startRow = createChartDataTable(workbook, sheet, sortedTasks, startRow);

            // Создаем диаграммы
            createSimpleStackedChart(workbook, (XSSFSheet) sheet, sortedTasks, dataStartRow, startRow);
            createExcelBarChart(workbook, (XSSFSheet) sheet, sortedTasks, dataStartRow, startRow + 25);
            createExcelLineChart(workbook, (XSSFSheet) sheet, sortedTasks, dataStartRow, startRow + 50);

        } catch (Exception e) {
            log.error("Ошибка при создании диаграмм Excel", e);
        }
    }

    /**
     * Простой вариант с использованием только high-level API
     */
    private void createSimpleStackedChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                          List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            Row chartTitleRow = sheet.createRow(chartRow);
            chartTitleRow.createCell(0).setCellValue("Распределение результатов (Stacked)");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    0, chartRow + 1, CHART_COL_SPAN, chartRow + CHART_ROW_SPAN);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Распределение результатов");

            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("№ задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Количество студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            // Просто используем STACKED группировку
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STACKED); // Просто STACKED
            barData.setVaryColors(true);

            // Используем абсолютные значения (не нормированные)
            CellRangeAddress labelRange = new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 1, 1));
            XDDFChartData.Series series1 = data.addSeries(xs, ys1);
            series1.setTitle("Полностью", null);

            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 2, 2));
            XDDFChartData.Series series2 = data.addSeries(xs, ys2);
            series2.setTitle("Частично", null);

            XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 3, 3));
            XDDFChartData.Series series3 = data.addSeries(xs, ys3);
            series3.setTitle("Не справилось", null);

            chart.plot(data);

            // Настраиваем цвета
            setSeriesColor(chart, 0, new Color(46, 204, 113));
            setSeriesColor(chart, 1, new Color(241, 196, 15));
            setSeriesColor(chart, 2, new Color(231, 76, 60));

            log.info("✅ Simple Stacked Chart создан");

        } catch (Exception e) {
            log.error("❌ Ошибка: {}", e.getMessage(), e);
        }
    }

    /**
     * Создает нормированную гистограмму с накоплением - ТОЧНО один столбец на задание
     */
    private void createTrueSingleColumnChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                             List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            Row chartTitleRow = sheet.createRow(chartRow);
            chartTitleRow.createCell(0).setCellValue("Нормированное распределение результатов по заданиям (%)");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    0, chartRow + 1, CHART_COL_SPAN, chartRow + CHART_ROW_SPAN);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Нормированное распределение результатов (%)");

            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("№ задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Доля студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            leftAxis.setMinimum(0.0);
            leftAxis.setMaximum(1.0);

            // Используем BAR, но с правильной конфигурацией
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STACKED);
            barData.setVaryColors(true);

            // Ключевая настройка: устанавливаем перекрытие столбцов на 100%
            // Это заставит столбцы накладываться друг на друга, создавая один общий столбец
            barData.setOverlap((byte) 100);

            // Создаем временные нормированные данные прямо здесь
            // Сначала создаем заголовок
            int tempDataRow = chartRow + 2;
            Row tempHeader = sheet.createRow(tempDataRow++);
            tempHeader.createCell(0).setCellValue("Задание");
            tempHeader.createCell(1).setCellValue("Полностью (%)");
            tempHeader.createCell(2).setCellValue("Частично (%)");
            tempHeader.createCell(3).setCellValue("Не справилось (%)");

            // Заполняем нормированные данные
            for (int i = 0; i < tasks.size(); i++) {
                TaskStatisticsDto task = tasks.get(i);
                Row row = sheet.createRow(tempDataRow++);

                row.createCell(0).setCellValue("№" + task.getTaskNumber());

                int total = task.getFullyCompletedCount() +
                        task.getPartiallyCompletedCount() +
                        task.getNotCompletedCount();

                if (total > 0) {
                    // Нормируем в проценты
                    Cell fullyCell = row.createCell(1);
                    double fullyPercent = (double) task.getFullyCompletedCount() / total;
                    fullyCell.setCellValue(fullyPercent);

                    Cell partiallyCell = row.createCell(2);
                    double partiallyPercent = (double) task.getPartiallyCompletedCount() / total;
                    partiallyCell.setCellValue(partiallyPercent);

                    Cell notCompletedCell = row.createCell(3);
                    double notCompletedPercent = (double) task.getNotCompletedCount() / total;
                    notCompletedCell.setCellValue(notCompletedPercent);
                } else {
                    row.createCell(1).setCellValue(0);
                    row.createCell(2).setCellValue(0);
                    row.createCell(3).setCellValue(0);
                }
            }

            // Теперь используем эти данные для диаграммы
            CellRangeAddress labelRange = new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 1, 1));
            XDDFChartData.Series series1 = data.addSeries(xs, ys1);
            series1.setTitle("Полностью", null);

            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 2, 2));
            XDDFChartData.Series series2 = data.addSeries(xs, ys2);
            series2.setTitle("Частично", null);

            XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(chartRow + 3, chartRow + 2 + tasks.size(), 3, 3));
            XDDFChartData.Series series3 = data.addSeries(xs, ys3);
            series3.setTitle("Не справилось", null);

            chart.plot(data);

            // Настраиваем цвета
            setSeriesColor(chart, 0, new Color(46, 204, 113)); // Зеленый
            setSeriesColor(chart, 1, new Color(241, 196, 15)); // Желтый
            setSeriesColor(chart, 2, new Color(231, 76, 60));  // Красный

            log.info("✅ Истинная нормированная гистограмма с одним столбцом создана");

        } catch (Exception e) {
            log.error("❌ Ошибка: {}", e.getMessage(), e);
        }
    }

    private int createChartDataTable(Workbook workbook, Sheet sheet,
                                     List<TaskStatisticsDto> tasks, int startRow) {

        Row headerRow = sheet.createRow(startRow++);
        String[] headers = {
                "Задание",
                "Полностью",
                "Частично",
                "Не справилось",
                "% выполнения"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        for (int i = 0; i < tasks.size(); i++) {
            TaskStatisticsDto task = tasks.get(i);
            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue("№" + task.getTaskNumber()); // Изменено здесь
            row.createCell(1).setCellValue(task.getFullyCompletedCount());
            row.createCell(2).setCellValue(task.getPartiallyCompletedCount());
            row.createCell(3).setCellValue(task.getNotCompletedCount());

            Cell percentCell = row.createCell(4);
            percentCell.setCellValue(task.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(createPercentStyle(workbook));
        }

        return startRow + 1;
    }

    /**
     * Создает Stacked Bar Chart в Excel
     */
    private void createExcelStackedBarChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                            List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            Row chartTitleRow = sheet.createRow(chartRow);
            chartTitleRow.createCell(0).setCellValue("Распределение результатов по заданиям");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    0, chartRow + 1, CHART_COL_SPAN, chartRow + CHART_ROW_SPAN);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Распределение результатов по заданиям");

            // Получаем легенду (она создается автоматически при создании диаграммы)
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Номер задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Количество студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            // Используем правильный тип диаграммы
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STACKED);  // Устанавливаем группировку в STACKED
            barData.setVaryColors(true);

            CellRangeAddress labelRange = new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            // Серия 1: Полностью
            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 1, 1));
            XDDFChartData.Series series1 = data.addSeries(xs, ys1);
            series1.setTitle("Полностью", null);

            // Серия 2: Частично
            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 2, 2));
            XDDFChartData.Series series2 = data.addSeries(xs, ys2);
            series2.setTitle("Частично", null);

            // Серия 3: Не справилось
            XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 3, 3));
            XDDFChartData.Series series3 = data.addSeries(xs, ys3);
            series3.setTitle("Не справилось", null);

            chart.plot(data);

            // Настраиваем цвета
            setSeriesColor(chart, 0, new Color(46, 204, 113)); // Зеленый
            setSeriesColor(chart, 1, new Color(241, 196, 15)); // Желтый
            setSeriesColor(chart, 2, new Color(231, 76, 60));  // Красный

            log.info("✅ Stacked Bar Chart создана успешно");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании Stacked Bar Chart: {}", e.getMessage(), e);
        }
    }

    /**
     * Создает Bar Chart в Excel для процентов выполнения
     */
    private void createExcelBarChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                     List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            Row chartTitleRow = sheet.createRow(chartRow);
            chartTitleRow.createCell(0).setCellValue("Процент выполнения заданий");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    0, chartRow + 1, CHART_COL_SPAN, chartRow + CHART_ROW_SPAN);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Процент выполнения заданий");

            // Получаем легенду
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Задание");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("% выполнения");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            leftAxis.setMinimum(0.0);
            leftAxis.setMaximum(1.0);

            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STANDARD);
            barData.setVaryColors(true);

            CellRangeAddress labelRange = new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 4, 4));

            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle("% выполнения", null);

            chart.plot(data);

            // Устанавливаем цвета в зависимости от процента выполнения
            // В этом случае все столбцы будут одного цвета (серия одна)
            // Если нужно разноцветные столбцы, нужно создавать отдельную серию для каждого столбца
            // или использовать другой подход
            setSeriesColor(chart, 0, new Color(52, 152, 219)); // Синий для всей серии

            log.info("✅ Bar Chart создана успешно");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании Bar Chart: {}", e.getMessage(), e);
        }
    }

    /**
     * Создает Line Chart в Excel для динамики выполнения
     */
    private void createExcelLineChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                      List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            Row chartTitleRow = sheet.createRow(chartRow);
            chartTitleRow.createCell(0).setCellValue("Динамика выполнения заданий");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    0, chartRow + 1, CHART_COL_SPAN, chartRow + CHART_ROW_SPAN);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Динамика выполнения заданий");

            // Получаем легенду
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Номер задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("% выполнения");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            leftAxis.setMinimum(0.0);
            leftAxis.setMaximum(1.0);

            XDDFChartData data = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
            XDDFLineChartData lineData = (XDDFLineChartData) data;

            CellRangeAddress labelRange = new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 4, 4));

            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle("% выполнения", null);

            // Настраиваем линию
            XDDFLineProperties lineProps = new XDDFLineProperties();
            lineProps.setWidth(2.0);

            // Создаем цвет через RGB
            byte[] rgbColor = new byte[3];
            rgbColor[0] = (byte) 52;   // R
            rgbColor[1] = (byte) 152;  // G
            rgbColor[2] = (byte) 219;  // B
            XDDFSolidFillProperties fill = new XDDFSolidFillProperties(new XDDFColorRgbBinary(rgbColor));
            lineProps.setFillProperties(fill);

            XDDFShapeProperties shapeProps = new XDDFShapeProperties();
            shapeProps.setLineProperties(lineProps);

            // Настройка маркеров (альтернативный способ)
            try {
                // Попытка создать маркеры через рефлексию (если класс доступен)
                Class<?> markerClass = Class.forName("org.apache.poi.xddf.usermodel.chart.XDDFMarker");
                Object marker = markerClass.getDeclaredConstructor().newInstance();
                markerClass.getMethod("setStyle", MarkerStyle.class).invoke(marker, MarkerStyle.CIRCLE);
                markerClass.getMethod("setSize", short.class).invoke(marker, (short) 6);
                lineData.getClass().getMethod("setMarker", markerClass).invoke(lineData, marker);
            } catch (Exception e) {
                log.warn("⚠️ Не удалось установить маркеры для линии: {}", e.getMessage());
                // Без маркеров тоже нормально
            }

            series.setShapeProperties(shapeProps);
            chart.plot(data);

            log.info("✅ Line Chart создана успешно");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании Line Chart: {}", e.getMessage(), e);
        }
    }

    /**
     * Устанавливает цвет для серии диаграммы
     */
    private void setSeriesColor(XSSFChart chart, int seriesIndex, Color color) {
        try {
            if (chart.getChartSeries() == null || chart.getChartSeries().isEmpty()) {
                return;
            }

            XDDFChartData data = chart.getChartSeries().get(0);
            if (seriesIndex >= data.getSeriesCount()) {
                return;
            }

            XDDFChartData.Series series = data.getSeries(seriesIndex);

            // Создаем цвет через RGB
            byte[] rgbColor = new byte[3];
            rgbColor[0] = (byte) color.getRed();
            rgbColor[1] = (byte) color.getGreen();
            rgbColor[2] = (byte) color.getBlue();
            XDDFSolidFillProperties fill = new XDDFSolidFillProperties(new XDDFColorRgbBinary(rgbColor));

            XDDFShapeProperties shapeProps = series.getShapeProperties();
            if (shapeProps == null) {
                shapeProps = new XDDFShapeProperties();
            }
            shapeProps.setFillProperties(fill);
            series.setShapeProperties(shapeProps);

        } catch (Exception e) {
            log.warn("⚠️ Не удалось установить цвет для серии {}: {}", seriesIndex, e.getMessage());
        }
    }

    /**
     * Возвращает цвет для процента выполнения
     */
    private Color getColorForPercentage(double percentage) {
        if (percentage >= 90) return new Color(39, 174, 96);    // Зеленый
        if (percentage >= 70) return new Color(46, 204, 113);   // Светло-зеленый
        if (percentage >= 50) return new Color(241, 196, 15);   // Желтый
        if (percentage >= 30) return new Color(230, 126, 34);   // Оранжевый
        return new Color(231, 76, 60);                         // Красный
    }

    private void createTeacherSummarySheet(XSSFWorkbook workbook, String teacherName,
                                           List<TestSummaryDto> teacherTests) {
        XSSFSheet sheet = workbook.createSheet("Сводка по тестам");

        XSSFRow titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Отчет по тестам учителя: " + teacherName);
        CellStyle titleStyle = workbook.createCellStyle();
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);
        titleStyle.setFont(titleFont);
        titleCell.setCellStyle(titleStyle);

        XSSFRow dateRow = sheet.createRow(1);
        dateRow.createCell(0).setCellValue(
                "Отчет сгенерирован: " +
                        LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE)
        );

        String[] headers = {
                "Предмет", "Класс", "Дата", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения"
        };

        createHeaderRow(sheet, workbook, headers, 3);

        int rowNum = 4;
        for (TestSummaryDto test : teacherTests) {
            XSSFRow row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(test.getSubject());
            row.createCell(1).setCellValue(test.getClassName());
            row.createCell(2).setCellValue(
                    test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            row.createCell(3).setCellValue(test.getTestType());

            setNumericCellValue(row, 4, test.getStudentsPresent());
            setNumericCellValue(row, 5, test.getStudentsAbsent());
            setNumericCellValue(row, 6, test.getClassSize());

            Cell attendanceCell = row.createCell(7);
            double attendance = test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0;
            attendanceCell.setCellValue(attendance);
            attendanceCell.setCellStyle(createPercentStyle(workbook));

            Cell avgScoreCell = row.createCell(8);
            avgScoreCell.setCellValue(test.getAverageScore() != null ?
                    test.getAverageScore() : 0.0);
            avgScoreCell.setCellStyle(createDecimalStyle(workbook));

            Cell successCell = row.createCell(9);
            double success = test.getSuccessPercentage() != null ?
                    test.getSuccessPercentage() / 100.0 : 0.0;
            successCell.setCellValue(success);
            successCell.setCellStyle(createPercentStyle(workbook));
        }

        addTeacherSummaryRow(sheet, rowNum, teacherTests, workbook);
        optimizeColumnWidths(sheet);
    }

    private void addTeacherSummaryRow(Sheet sheet, int rowNum,
                                      List<TestSummaryDto> tests, Workbook workbook) {
        if (tests.isEmpty()) return;

        Row summaryRow = sheet.createRow(rowNum + 1);
        CellStyle summaryStyle = createSummaryStyle((XSSFWorkbook) workbook);

        summaryRow.createCell(0).setCellValue("Средние показатели:");
        summaryRow.getCell(0).setCellStyle(summaryStyle);

        double avgPresent = tests.stream()
                .filter(t -> t.getStudentsPresent() != null)
                .mapToInt(TestSummaryDto::getStudentsPresent)
                .average().orElse(0);

        double avgAbsent = tests.stream()
                .filter(t -> t.getStudentsAbsent() != null)
                .mapToInt(TestSummaryDto::getStudentsAbsent)
                .average().orElse(0);

        double avgClassSize = tests.stream()
                .filter(t -> t.getClassSize() != null)
                .mapToInt(TestSummaryDto::getClassSize)
                .average().orElse(0);

        double avgAttendance = tests.stream()
                .filter(t -> t.getAttendancePercentage() != null)
                .mapToDouble(t -> t.getAttendancePercentage() / 100.0)
                .average().orElse(0);

        double avgScore = tests.stream()
                .filter(t -> t.getAverageScore() != null)
                .mapToDouble(TestSummaryDto::getAverageScore)
                .average().orElse(0);

        double avgSuccess = tests.stream()
                .filter(t -> t.getSuccessPercentage() != null)
                .mapToDouble(t -> t.getSuccessPercentage() / 100.0)
                .average().orElse(0);

        Cell avgPresentCell = summaryRow.createCell(4);
        avgPresentCell.setCellValue(avgPresent);
        avgPresentCell.setCellStyle(summaryStyle);

        Cell avgAbsentCell = summaryRow.createCell(5);
        avgAbsentCell.setCellValue(avgAbsent);
        avgAbsentCell.setCellStyle(summaryStyle);

        Cell avgClassSizeCell = summaryRow.createCell(6);
        avgClassSizeCell.setCellValue(avgClassSize);
        avgClassSizeCell.setCellStyle(summaryStyle);

        Cell avgAttendanceCell = summaryRow.createCell(7);
        avgAttendanceCell.setCellValue(avgAttendance);
        avgAttendanceCell.setCellStyle(createPercentStyle(workbook));

        Cell avgScoreCell = summaryRow.createCell(8);
        avgScoreCell.setCellValue(avgScore);
        avgScoreCell.setCellStyle(createDecimalStyle(workbook));

        Cell avgSuccessCell = summaryRow.createCell(9);
        avgSuccessCell.setCellValue(avgSuccess);
        avgSuccessCell.setCellStyle(createPercentStyle(workbook));
    }

    private void setCellValue(Row row, int cellIndex, Object value) {
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

    private void setNumericCellValue(Row row, int cellIndex, Integer value) {
        Cell cell = row.createCell(cellIndex);
        if (value != null) {
            cell.setCellValue(value);
        } else {
            cell.setCellValue(0);
        }
    }

    private void setNumericCellValue(Row row, int cellIndex, Double value) {
        Cell cell = row.createCell(cellIndex);
        if (value != null) {
            cell.setCellValue(value);
        } else {
            cell.setCellValue(0);
        }
    }

    private void createTeacherTestSheet(XSSFWorkbook workbook, TestSummaryDto test,
                                        String teacherName, int sheetIndex) {
        String sheetName = String.format("%s_%s_%s",
                test.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", "").substring(0,
                        Math.min(12, test.getSubject().length())),
                test.getClassName(),
                test.getTestDate().format(DateTimeFormatter.ofPattern("ddMM")));

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        String finalSheetName = sheetName;
        int counter = 1;
        while (workbook.getSheet(finalSheetName) != null) {
            finalSheetName = String.format("%s_%d", sheetName, counter++);
            if (finalSheetName.length() > 31) {
                finalSheetName = finalSheetName.substring(0, 31);
            }
        }

        XSSFSheet sheet = workbook.createSheet(finalSheetName);

        int rowNum = 0;
        XSSFRow titleRow = sheet.createRow(rowNum++);
        titleRow.createCell(0).setCellValue(
                String.format("Краткая статистика: %s - %s",
                        test.getSubject(), test.getClassName())
        );
        titleRow.getCell(0).setCellStyle(createTitleStyle(workbook));

        rowNum++;
        createTestStatisticsTable(sheet, workbook, test, rowNum);
        optimizeColumnWidths(sheet);
    }

    private void createTestStatisticsTable(Sheet sheet, Workbook workbook,
                                           TestSummaryDto test, int startRow) {
        Row headerRow = sheet.createRow(startRow++);
        headerRow.createCell(0).setCellValue("Показатель");
        headerRow.createCell(1).setCellValue("Значение");
        headerRow.getCell(0).setCellStyle(createTableHeaderStyle(workbook));
        headerRow.getCell(1).setCellStyle(createTableHeaderStyle(workbook));

        String[][] stats = {
                {"Предмет", test.getSubject()},
                {"Класс", test.getClassName()},
                {"Учитель", test.getTeacher()},
                {"Дата теста", test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE)},
                {"Тип работы", test.getTestType()},
                {"Всего учеников", String.valueOf(test.getClassSize())},
                {"Присутствовало", String.valueOf(test.getStudentsPresent())},
                {"Отсутствовало", String.valueOf(test.getStudentsAbsent())},
                {"% присутствия", String.format("%.1f%%", test.getAttendancePercentage())},
                {"Количество заданий", String.valueOf(test.getTaskCount())},
                {"Макс. балл", String.valueOf(test.getMaxTotalScore())},
                {"Средний балл", String.format("%.2f", test.getAverageScore())},
                {"% выполнения", String.format("%.1f%%", test.getSuccessPercentage())},
                {"Файл отчета", test.getFileName()}
        };

        for (String[] rowData : stats) {
            Row row = sheet.createRow(startRow++);
            row.createCell(0).setCellValue(rowData[0]);
            row.createCell(1).setCellValue(rowData[1]);
        }
    }

    private void createSummarySheetHeader(Sheet sheet, Workbook workbook) {
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("СВОДНЫЙ ОТЧЕТ ПО ВСЕМ ТЕСТАМ");
        CellStyle titleStyle = createTitleStyle(workbook);
        titleCell.setCellStyle(titleStyle);

        Row dateRow = sheet.createRow(1);
        dateRow.createCell(0).setCellValue(
                "Дата формирования: " +
                        LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE)
        );

        String[] headers = {
                "ID", "Предмет", "Класс", "Учитель", "Дата теста", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения", "Кол-во заданий", "Макс. балл"
        };

        createHeaderRow(sheet, workbook, headers, 3);
    }

    private void createHeaderRow(Sheet sheet, Workbook workbook, String[] headers, int rowNum) {
        Row headerRow = sheet.createRow(rowNum);
        CellStyle headerStyle = createTableHeaderStyle(workbook);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void fillSummaryData(Sheet sheet, Workbook workbook, List<TestSummaryDto> tests) {
        int rowNum = 4;
        int testId = 1;

        for (TestSummaryDto test : tests) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(testId++);
            row.createCell(1).setCellValue(test.getSubject());
            row.createCell(2).setCellValue(test.getClassName());
            row.createCell(3).setCellValue(test.getTeacher());
            row.createCell(4).setCellValue(
                    test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            row.createCell(5).setCellValue(test.getTestType());

            setNumericCellValue(row, 6, test.getStudentsPresent());
            setNumericCellValue(row, 7, test.getStudentsAbsent());
            setNumericCellValue(row, 8, test.getClassSize());

            Cell attendanceCell = row.createCell(9);
            double attendance = test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0;
            attendanceCell.setCellValue(attendance);
            attendanceCell.setCellStyle(createPercentStyle(workbook));

            Cell avgScoreCell = row.createCell(10);
            avgScoreCell.setCellValue(test.getAverageScore() != null ?
                    test.getAverageScore() : 0.0);
            avgScoreCell.setCellStyle(createDecimalStyle(workbook));

            Cell successCell = row.createCell(11);
            double success = test.getSuccessPercentage() != null ?
                    test.getSuccessPercentage() / 100.0 : 0.0;
            successCell.setCellValue(success);
            successCell.setCellStyle(createPercentStyle(workbook));

            setNumericCellValue(row, 12, test.getTaskCount());
            setNumericCellValue(row, 13, test.getMaxTotalScore());
        }
    }

    private void addSummaryStatisticsRow(Sheet sheet, List<TestSummaryDto> tests) {
        if (tests.isEmpty()) return;

        int lastRow = sheet.getLastRowNum() + 2;
        Row summaryRow = sheet.createRow(lastRow);

        summaryRow.createCell(0).setCellValue("ИТОГО:");
        CellStyle summaryStyle = createSummaryStyle((XSSFWorkbook) sheet.getWorkbook());
        summaryRow.getCell(0).setCellStyle(summaryStyle);

        int totalTests = tests.size();
        summaryRow.createCell(1).setCellValue("Тестов: " + totalTests);

        double avgPresent = tests.stream()
                .filter(t -> t.getStudentsPresent() != null)
                .mapToInt(TestSummaryDto::getStudentsPresent)
                .average().orElse(0);

        double avgAbsent = tests.stream()
                .filter(t -> t.getStudentsAbsent() != null)
                .mapToInt(TestSummaryDto::getStudentsAbsent)
                .average().orElse(0);

        double avgClassSize = tests.stream()
                .filter(t -> t.getClassSize() != null)
                .mapToInt(TestSummaryDto::getClassSize)
                .average().orElse(0);

        double avgAttendance = tests.stream()
                .filter(t -> t.getAttendancePercentage() != null)
                .mapToDouble(t -> t.getAttendancePercentage() / 100.0)
                .average().orElse(0);

        double avgScore = tests.stream()
                .filter(t -> t.getAverageScore() != null)
                .mapToDouble(TestSummaryDto::getAverageScore)
                .average().orElse(0);

        double avgSuccess = tests.stream()
                .filter(t -> t.getSuccessPercentage() != null)
                .mapToDouble(t -> t.getSuccessPercentage() / 100.0)
                .average().orElse(0);

        Cell avgPresentCell = summaryRow.createCell(6);
        avgPresentCell.setCellValue(String.format("%.1f", avgPresent));

        Cell avgAbsentCell = summaryRow.createCell(7);
        avgAbsentCell.setCellValue(String.format("%.1f", avgAbsent));

        Cell avgClassSizeCell = summaryRow.createCell(8);
        avgClassSizeCell.setCellValue(String.format("%.1f", avgClassSize));

        Cell avgAttendanceCell = summaryRow.createCell(9);
        avgAttendanceCell.setCellValue(avgAttendance);
        avgAttendanceCell.setCellStyle(createPercentStyle(sheet.getWorkbook()));

        Cell avgScoreCell = summaryRow.createCell(10);
        avgScoreCell.setCellValue(String.format("%.2f", avgScore));

        Cell avgSuccessCell = summaryRow.createCell(11);
        avgSuccessCell.setCellValue(avgSuccess);
        avgSuccessCell.setCellStyle(createPercentStyle(sheet.getWorkbook()));
    }

    private Path createReportsFolder() throws IOException {
        Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
        if (!Files.exists(reportsPath)) {
            Files.createDirectories(reportsPath);
        }
        return reportsPath;
    }

    private File saveWorkbook(Workbook workbook, Path folderPath, String fileName) throws IOException {
        Path filePath = folderPath.resolve(fileName);
        try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
            workbook.write(fos);
        }
        log.info("Отчет сохранен: {}", filePath);
        return filePath.toFile();
    }

    private void optimizeColumnWidths(Sheet sheet) {
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
            int currentWidth = sheet.getColumnWidth(i);
            if (currentWidth < MIN_COLUMN_WIDTH) {
                sheet.setColumnWidth(i, MIN_COLUMN_WIDTH);
            } else if (currentWidth > MAX_COLUMN_WIDTH) {
                sheet.setColumnWidth(i, MAX_COLUMN_WIDTH);
            }
        }
    }

    private CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createSubtitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

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

    private CellStyle createCenteredStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

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

    private CellStyle createDecimalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createSummaryStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }
}