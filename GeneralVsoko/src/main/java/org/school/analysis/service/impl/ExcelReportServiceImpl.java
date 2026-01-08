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
    private static final int CHART_COL_SPAN = 10;
    private static final int CHART_ROW_SPAN = 15;
    private static final int MIN_COLUMN_WIDTH = 8 * 256;  // Уменьшил минимальную ширину
    private static final int MAX_COLUMN_WIDTH = 40 * 256; // Уменьшил максимальную ширину

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
                optimizeColumnWidths(sheet, 14); // Указываем количество колонок

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

            // Определяем количество заданий
            int maxTasks = determineMaxTaskCount(studentResults, taskStatistics);

            currentRow = createStudentResultsSection(sheet, workbook, studentResults, currentRow, maxTasks);
            currentRow = createTaskAnalysisSection(sheet, workbook, taskStatistics, currentRow, maxTasks);

            if (taskStatistics != null && !taskStatistics.isEmpty() &&
                    studentResults != null && !studentResults.isEmpty()) {
                createExcelChartsSection(workbook, sheet, testSummary, taskStatistics, studentResults, currentRow);
            }

            // Оптимизируем ширину колонок с учетом количества заданий
            optimizeAllColumns(sheet, maxTasks);

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета на одном листе", e);
            XSSFRow errorRow = sheet.createRow(currentRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании отчета: " + e.getMessage());
        }
    }

    /**
     * Определяет максимальное количество заданий
     */
    private int determineMaxTaskCount(List<StudentDetailedResultDto> studentResults,
                                      Map<Integer, TaskStatisticsDto> taskStatistics) {
        int maxTasks = 0;

        // Сначала проверяем статистику
        if (taskStatistics != null && !taskStatistics.isEmpty()) {
            maxTasks = taskStatistics.keySet().stream()
                    .mapToInt(Integer::intValue)
                    .max()
                    .orElse(0);
        }

        // Затем проверяем студентов
        if (studentResults != null && !studentResults.isEmpty()) {
            for (StudentDetailedResultDto student : studentResults) {
                if (student.getTaskScores() != null) {
                    int studentMaxTask = student.getTaskScores().keySet().stream()
                            .mapToInt(Integer::intValue)
                            .max()
                            .orElse(0);
                    maxTasks = Math.max(maxTasks, studentMaxTask);
                }
            }
        }

        // Если всё еще 0, используем значение по умолчанию
        if (maxTasks == 0) {
            maxTasks = 10;
        }

        // Гарантируем минимум 10 заданий для корректного отображения
        return Math.max(maxTasks, 10);
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

        // Объединяем ячейки для заголовка (A1:Y1 для 25 колонок)
        sheet.addMergedRegion(new CellRangeAddress(
                startRow - 1, startRow - 1, 0, 24
        ));

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
                                            int startRow, int maxTasks) {
        if (studentResults == null || studentResults.isEmpty()) {
            return startRow;
        }

        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("РЕЗУЛЬТАТЫ СТУДЕНТОВ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        List<String> headers = new ArrayList<>(Arrays.asList(
                "№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"
        ));

        for (int i = 1; i <= maxTasks; i++) {
            headers.add("№" + i);
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
                                          int startRow, int maxTasks) {
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

        // Создаем гарантированный список заданий от 1 до maxTasks
        for (int taskNum = 1; taskNum <= maxTasks; taskNum++) {
            TaskStatisticsDto stats = taskStatistics.get(taskNum);
            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue("№" + taskNum);

            if (stats != null) {
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
                } else {
                    row.createCell(6).setCellValue("-");
                }
            } else {
                // Заполняем значения по умолчанию для отсутствующих заданий
                row.createCell(1).setCellValue(1);
                row.createCell(2).setCellValue(0);
                row.createCell(3).setCellValue(0);
                row.createCell(4).setCellValue(0);

                Cell percentCell = row.createCell(5);
                percentCell.setCellValue(0);
                percentCell.setCellStyle(createPercentStyle(workbook));

                row.createCell(6).setCellValue("-");
            }
        }

        startRow++;
        return startRow;
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

            startRow++; // Пустая строка после заголовка

            if (taskStatistics == null || taskStatistics.isEmpty()) {
                return;
            }

            List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                    .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                    .collect(Collectors.toList());

            int dataStartRow = startRow;
            startRow = createChartDataTable(workbook, sheet, sortedTasks, startRow);

            int chartStartRow = startRow + 2; // Отступ после таблицы данных

            // Создаем диаграммы последовательно
            createSimpleStackedChart(workbook, (XSSFSheet) sheet, sortedTasks,
                    dataStartRow, chartStartRow);

            // Вторая диаграмма смещена на CHART_ROW_SPAN + 2 строки
            createExcelBarChart(workbook, (XSSFSheet) sheet, sortedTasks,
                    dataStartRow, chartStartRow + CHART_ROW_SPAN + 2);

            // Третья диаграмма смещена еще больше
            createExcelLineChart(workbook, (XSSFSheet) sheet, sortedTasks,
                    dataStartRow, chartStartRow + 2 * CHART_ROW_SPAN + 4);

        } catch (Exception e) {
            log.error("Ошибка при создании диаграмм Excel", e);
        }
    }

    /**
     * Улучшенный метод создания таблицы данных для графиков
     */
    /**
     * Создает таблицу данных для графиков с проверкой позиции
     */
    private int createChartDataTable(Workbook workbook, Sheet sheet,
                                     List<TaskStatisticsDto> tasks, int startRow) {

        // Проверяем, не перекрывает ли таблица существующие данные
        int lastRowWithData = sheet.getLastRowNum();
        if (startRow <= lastRowWithData) {
            startRow = lastRowWithData + 3; // Добавляем отступ
        }

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

            row.createCell(0).setCellValue("№" + task.getTaskNumber());
            row.createCell(1).setCellValue(task.getFullyCompletedCount());
            row.createCell(2).setCellValue(task.getPartiallyCompletedCount());
            row.createCell(3).setCellValue(task.getNotCompletedCount());

            Cell percentCell = row.createCell(4);
            percentCell.setCellValue(task.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(createPercentStyle(workbook));
        }

        return startRow + 2; // Добавляем дополнительный отступ
    }

    /**
     * Исправленный метод создания Stacked диаграммы
     */
    private void createSimpleStackedChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                          List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            // Добавляем пустую строку перед графиком
            sheet.createRow(chartRow);

            Row chartTitleRow = sheet.createRow(chartRow + 1);
            chartTitleRow.createCell(0).setCellValue("Распределение результатов (Stacked)");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // Сдвигаем график вправо (col1 вместо col0) и увеличиваем ширину
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    1, chartRow + 2, CHART_COL_SPAN + 3, chartRow + CHART_ROW_SPAN + 2);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Распределение результатов");

            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("№ задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Количество студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STACKED);
            barData.setVaryColors(true);

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
            setSeriesColor(chart, 0, new Color(0, 20, 236));   // Синий
            setSeriesColor(chart, 1, new Color(241, 120, 0));  // Оранжевый
            setSeriesColor(chart, 2, new Color(231, 50, 60));  // Красный

            log.info("✅ Simple Stacked Chart создан");

        } catch (Exception e) {
            log.error("❌ Ошибка: {}", e.getMessage(), e);
        }
    }

    /**
     * Исправленный метод создания Bar диаграммы
     */
    private void createExcelBarChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                     List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            // Добавляем пустую строку перед графиком
            sheet.createRow(chartRow);

            Row chartTitleRow = sheet.createRow(chartRow + 1);
            chartTitleRow.createCell(0).setCellValue("Процент выполнения заданий");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // Сдвигаем график вправо
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    1, chartRow + 2, CHART_COL_SPAN + 3, chartRow + CHART_ROW_SPAN + 2);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Процент выполнения заданий");

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
            setSeriesColor(chart, 0, new Color(52, 152, 219));

            log.info("✅ Bar Chart создана успешно");

        } catch (Exception e) {
            log.error("❌ Ошибка при создании Bar Chart: {}", e.getMessage(), e);
        }
    }

    /**
     * Исправленный метод создания Line диаграммы
     */
    private void createExcelLineChart(XSSFWorkbook workbook, XSSFSheet sheet,
                                      List<TaskStatisticsDto> tasks, int dataStartRow, int chartRow) {

        try {
            // Добавляем пустую строку перед графиком
            sheet.createRow(chartRow);

            Row chartTitleRow = sheet.createRow(chartRow + 1);
            chartTitleRow.createCell(0).setCellValue("Динамика выполнения заданий");
            chartTitleRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // Сдвигаем график вправо
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    1, chartRow + 2, CHART_COL_SPAN + 3, chartRow + CHART_ROW_SPAN + 2);

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Динамика выполнения заданий");

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

            XDDFLineProperties lineProps = new XDDFLineProperties();
            lineProps.setWidth(2.0);

            byte[] rgbColor = new byte[3];
            rgbColor[0] = (byte) 52;   // R
            rgbColor[1] = (byte) 152;  // G
            rgbColor[2] = (byte) 219;  // B
            XDDFSolidFillProperties fill = new XDDFSolidFillProperties(new XDDFColorRgbBinary(rgbColor));
            lineProps.setFillProperties(fill);

            XDDFShapeProperties shapeProps = new XDDFShapeProperties();
            shapeProps.setLineProperties(lineProps);

            try {
                Class<?> markerClass = Class.forName("org.apache.poi.xddf.usermodel.chart.XDDFMarker");
                Object marker = markerClass.getDeclaredConstructor().newInstance();
                markerClass.getMethod("setStyle", MarkerStyle.class).invoke(marker, MarkerStyle.CIRCLE);
                markerClass.getMethod("setSize", short.class).invoke(marker, (short) 6);
                lineData.getClass().getMethod("setMarker", markerClass).invoke(lineData, marker);
            } catch (Exception e) {
                log.warn("⚠️ Не удалось установить маркеры для линии: {}", e.getMessage());
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

    private void createTeacherSummarySheet(XSSFWorkbook workbook, String teacherName,
                                           List<TestSummaryDto> teacherTests) {
        XSSFSheet sheet = workbook.createSheet("Сводка по тестам");

        XSSFRow titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Отчет по тестам учителя: " + teacherName);
        CellStyle titleStyle = createTitleStyle(workbook);
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
        optimizeColumnWidths(sheet, headers.length);
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
        optimizeColumnWidths(sheet, 2);
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
                "Учебный год", "Предмет", "Класс", "Учитель", "Дата теста", "Тип",
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

            row.createCell(0).setCellValue(test.getACADEMIC_YEAR());
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

    /**
     * Улучшенная оптимизация ширины колонок
     */
    private void optimizeColumnWidths(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
            int currentWidth = sheet.getColumnWidth(i);

            // Устанавливаем разумные ограничения
            if (currentWidth < MIN_COLUMN_WIDTH) {
                sheet.setColumnWidth(i, MIN_COLUMN_WIDTH);
            } else if (currentWidth > MAX_COLUMN_WIDTH) {
                sheet.setColumnWidth(i, MAX_COLUMN_WIDTH);
            }
        }
    }

    /**
     * Улучшенная оптимизация ширины для детальных отчетов
     */
    private void optimizeAllColumns(Sheet sheet, int maxTasks) {
        // Базовые ширины для основных колонок
        int[] baseWidths = {
                1500,   // №
                4500,   // ФИО (уменьшено)
                2000,   // Присутствие
                1500,   // Вариант
                1500,   // Общий балл
                1500,   // % выполнения
        };

        // Ширины для секции анализа заданий (G, H, I...)
        for (int i = 0; i < baseWidths.length; i++) {
            sheet.setColumnWidth(i, baseWidths[i]);
        }

        // Колонки с результатами заданий (начиная с 7-й колонки)
        int taskResultWidth = 800; // Уменьшенная ширина для баллов заданий
        for (int i = 0; i < maxTasks; i++) {
            int colIndex = 6 + i; // Начинаем с колонки G (индекс 6)
            sheet.setColumnWidth(colIndex, taskResultWidth);

            // Устанавливаем стиль центрирования
            CellStyle centeredStyle = createCenteredStyle(sheet.getWorkbook());
            // Применяем стиль к заголовкам заданий
            if (sheet.getRow(10) != null && sheet.getRow(10).getCell(colIndex) != null) {
                sheet.getRow(10).getCell(colIndex).setCellStyle(centeredStyle);
            }
        }

        // Колонки для анализа заданий (начиная с 6 + maxTasks)
        int analysisStartCol = 6 + maxTasks;
        int[] analysisWidths = {1500, 1500, 1500, 1500, 1500, 1500, 3000}; // 7 колонок

        for (int i = 0; i < analysisWidths.length; i++) {
            int colIndex = analysisStartCol + i;
            if (colIndex < 256) { // Максимум 256 колонок в Excel
                sheet.setColumnWidth(colIndex, analysisWidths[i]);
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