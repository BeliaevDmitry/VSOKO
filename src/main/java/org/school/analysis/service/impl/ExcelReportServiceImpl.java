package org.school.analysis.service.impl;

import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.axis.ValueAxis;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.util.ModernGradientBarPainter;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.labels.CategoryItemLabelGenerator;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.plot.*;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.BoxAndWhiskerRenderer;
import org.jfree.chart.renderer.category.GradientBarPainter;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.chart.ui.TextAnchor;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.statistics.DefaultBoxAndWhiskerCategoryDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.ExcelReportService;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
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
    private static final int CHART_WIDTH = 600;
    private static final int CHART_HEIGHT = 400;
    private static final int IMAGE_COL_WIDTH = 30;
    private static final int IMAGE_ROW_HEIGHT = 20 * 20; // 20 строк по 20 пунктов

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
                autoSizeColumns(sheet, 14);

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
                // Используем ту же логику, что и для отчета учителя
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

                // Создаем отдельные листы для каждого теста
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

                // 1. Сводный лист по всем тестам
                createTeacherSummarySheet(workbook, teacherName, teacherTests);

                // 2. Для каждого теста создаем ПОЛНЫЙ детальный отчет на отдельном листе
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

    /**
     * Создает полный детальный отчет по тесту для отчета учителя
     */
    private void createCompleteTestDetailSheetForTeacher(XSSFWorkbook workbook,
                                                         TeacherTestDetailDto testDetail,
                                                         String teacherName,
                                                         int sheetIndex) {

        TestSummaryDto testSummary = testDetail.getTestSummary();
        if (testSummary == null) {
            log.warn("Пропускаем тест без основных данных");
            return;
        }

        // Генерируем имя листа
        String sheetName = String.format("%s_%s_%s",
                testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", "").substring(0,
                        Math.min(12, testSummary.getSubject().length())),
                testSummary.getClassName(),
                testSummary.getTestDate().format(DateTimeFormatter.ofPattern("ddMM")));

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        // Проверяем уникальность имени
        String finalSheetName = sheetName;
        int counter = 1;
        while (workbook.getSheet(finalSheetName) != null) {
            finalSheetName = String.format("%s_%d", sheetName, counter++);
            if (finalSheetName.length() > 31) {
                finalSheetName = finalSheetName.substring(0, 31);
            }
        }

        try {
            // Создаем полный детальный отчет на одном листе
            createSingleSheetDetailReport(workbook, testSummary,
                    testDetail.getStudentResults(),
                    testDetail.getTaskStatistics(),
                    finalSheetName);

        } catch (Exception e) {
            log.error("Ошибка при создании детального листа для теста {}: {}",
                    testSummary.getFileName(), e.getMessage(), e);
        }
    }

    /**
     * Создает детальный отчет по тесту на одном листе (переименованная версия)
     */
    private void createSingleSheetDetailReport(XSSFWorkbook workbook,
                                               TestSummaryDto testSummary,
                                               List<StudentDetailedResultDto> studentResults,
                                               Map<Integer, TaskStatisticsDto> taskStatistics,
                                               String sheetName) {

        XSSFSheet sheet = workbook.createSheet(sheetName);

        int currentRow = 0;

        try {
            // 1. ЗАГОЛОВОК И ОСНОВНАЯ ИНФОРМАЦИЯ
            currentRow = createReportHeader(sheet, workbook, testSummary, currentRow);

            // 2. ОБЩАЯ СТАТИСТИКА ТЕСТА
            currentRow = createTestStatisticsSection(sheet, workbook, testSummary, currentRow);

            // 3. РЕЗУЛЬТАТЫ СТУДЕНТОВ
            currentRow = createStudentResultsSection(sheet, workbook, studentResults, currentRow);

            // 4. АНАЛИЗ ПО ЗАДАНИЯМ
            currentRow = createTaskAnalysisSection(sheet, workbook, taskStatistics, currentRow);

            // 5. ГРАФИКИ АНАЛИЗА (если есть данные)
            if (taskStatistics != null && !taskStatistics.isEmpty() &&
                    studentResults != null && !studentResults.isEmpty()) {
                createChartsSection(workbook, sheet, testSummary, taskStatistics, studentResults, currentRow);
            }

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета на одном листе", e);
            XSSFRow errorRow = sheet.createRow(currentRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании отчета: " + e.getMessage());
        }
    }

    /**
     * Создает заголовок отчета
     */
    private int createReportHeader(Sheet sheet, Workbook workbook,
                                   TestSummaryDto testSummary, int startRow) {
        // Основной заголовок
        Row titleRow = sheet.createRow(startRow++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(
                String.format("ДЕТАЛЬНЫЙ ОТЧЕТ ПО ТЕСТУ: %s - %s",
                        testSummary.getSubject(), testSummary.getClassName())
        );
        CellStyle titleStyle = createTitleStyle(workbook);
        titleCell.setCellStyle(titleStyle);

        // Подзаголовок с датой
        Row subtitleRow = sheet.createRow(startRow++);
        Cell subtitleCell = subtitleRow.createCell(0);
        subtitleCell.setCellValue(
                String.format("Дата проведения: %s | Тип работы: %s",
                        testSummary.getTestDate().format(DateTimeFormatters.DISPLAY_DATE),
                        testSummary.getTestType())
        );
        CellStyle subtitleStyle = createSubtitleStyle(workbook);
        subtitleCell.setCellStyle(subtitleStyle);

        // Учитель
        Row teacherRow = sheet.createRow(startRow++);
        teacherRow.createCell(0).setCellValue("Учитель: " + testSummary.getTeacher());

        // Файл отчета
        Row fileRow = sheet.createRow(startRow++);
        fileRow.createCell(0).setCellValue("Файл отчета: " + testSummary.getFileName());

        // Пустая строка
        startRow++;

        return startRow;
    }
    /**
     * Создает секцию с общей статистикой теста
     */
    private int createTestStatisticsSection(Sheet sheet, Workbook workbook,
                                            TestSummaryDto testSummary, int startRow) {
        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("ОБЩАЯ СТАТИСТИКА ТЕСТА");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Создаем таблицу с двумя колонками
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

        // Пустая строка
        startRow++;

        return startRow;
    }

    /**
     * Создает секцию с результатами студентов
     */
    private int createStudentResultsSection(Sheet sheet, Workbook workbook,
                                            List<StudentDetailedResultDto> studentResults,
                                            int startRow) {
        if (studentResults == null || studentResults.isEmpty()) {
            return startRow;
        }

        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("РЕЗУЛЬТАТЫ СТУДЕНТОВ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Определяем максимальное количество заданий
        int maxTasks = studentResults.stream()
                .map(StudentDetailedResultDto::getTaskScores)
                .filter(Objects::nonNull)
                .mapToInt(Map::size)
                .max()
                .orElse(0);

        // Заголовки таблицы
        List<String> headers = new ArrayList<>(Arrays.asList(
                "№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"
        ));

        for (int i = 1; i <= maxTasks; i++) {
            headers.add("З" + i);
        }

        Row headerRow = sheet.createRow(startRow++);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        // Данные студентов
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

            // Баллы по заданиям
            Map<Integer, Integer> taskScores = student.getTaskScores();
            for (int taskNum = 1; taskNum <= maxTasks; taskNum++) {
                Integer score = taskScores != null ? taskScores.get(taskNum) : null;
                Cell scoreCell = row.createCell(5 + taskNum);
                scoreCell.setCellValue(score != null ? score : 0);
                scoreCell.setCellStyle(createCenteredStyle(workbook));
            }
        }

        // Пустая строка
        startRow++;

        return startRow;
    }
    /**
     * Создает секцию с анализом заданий
     */
    private int createTaskAnalysisSection(Sheet sheet, Workbook workbook,
                                          Map<Integer, TaskStatisticsDto> taskStatistics,
                                          int startRow) {
        if (taskStatistics == null || taskStatistics.isEmpty()) {
            return startRow;
        }

        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("АНАЛИЗ ПО ЗАДАНИЯМ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Заголовки таблицы
        String[] headers = {
                "Задание", "Макс. балл", "Полностью", "Частично", "Не справилось",
                "% выполнения", "Распределение баллов"
        };

        Row headerRow = sheet.createRow(startRow++);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        // Данные по заданиям
        for (TaskStatisticsDto stats : taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList())) {
            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue(stats.getTaskNumber());
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

        // Пустая строка
        startRow++;

        return startRow;
    }

    /**
     * Создает секцию с графиками
     */
    private void createChartsSection(XSSFWorkbook workbook, Sheet sheet,
                                     TestSummaryDto testSummary,
                                     Map<Integer, TaskStatisticsDto> taskStatistics,
                                     List<StudentDetailedResultDto> studentResults,
                                     int startRow) {
        try {
            // Заголовок секции
            Row sectionHeader = sheet.createRow(startRow++);
            sectionHeader.createCell(0).setCellValue("ГРАФИЧЕСКИЙ АНАЛИЗ");
            sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

            // Создаем графики
            JFreeChart chart1 = createEnhancedStackedBarChart(taskStatistics);
            JFreeChart chart2 = createFixedHorizontalBarChart(taskStatistics);
            JFreeChart chart3 = createTrendLineChart(taskStatistics);

            // Вставляем графики как изображения
            int chartRow = startRow;
            chartRow = insertChartAsImage(workbook, sheet, chart1,
                    "Распределение результатов по заданиям", chartRow);

            chartRow = insertChartAsImage(workbook, sheet, chart2,
                    "Процент выполнения заданий", chartRow + 2);

            insertChartAsImage(workbook, sheet, chart3,
                    "Динамика выполнения заданий с линией тренда", chartRow + 2);

        } catch (Exception e) {
            log.error("Ошибка при создании графиков в детальном отчете", e);
        }
    }

    private void createTeacherSummarySheet(XSSFWorkbook workbook, String teacherName,
                                           List<TestSummaryDto> teacherTests) {
        XSSFSheet sheet = workbook.createSheet("Сводка по тестам");

        // Заголовок отчета
        XSSFRow titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("Отчет по тестам учителя: " + teacherName);
        CellStyle titleStyle = workbook.createCellStyle();
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);
        titleStyle.setFont(titleFont);
        titleCell.setCellStyle(titleStyle);

        // Дата генерации отчета
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

        // Данные
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

        // Добавляем строку с итогами
        addTeacherSummaryRow(sheet, rowNum, teacherTests, workbook);

        autoSizeColumns(sheet, headers.length);
    }

    /**
     * Добавляет строку с итогами в отчет учителя
     */
    private void addTeacherSummaryRow(Sheet sheet, int rowNum,
                                      List<TestSummaryDto> tests, Workbook workbook) {
        if (tests.isEmpty()) return;

        Row summaryRow = sheet.createRow(rowNum + 1);
        CellStyle summaryStyle = createSummaryStyle((XSSFWorkbook) workbook);

        summaryRow.createCell(0).setCellValue("Средние показатели:");
        summaryRow.getCell(0).setCellStyle(summaryStyle);

        // Рассчитываем средние значения
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

        // Заполняем средние значения
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

    /**
     * Вставляет диаграмму как изображение в лист
     */
    private int insertChartAsImage(XSSFWorkbook workbook, Sheet sheet,
                                   JFreeChart chart, String chartTitle, int startRow) throws IOException {
        // Заголовок диаграммы
        Row titleRow = sheet.createRow(startRow);
        titleRow.createCell(0).setCellValue(chartTitle);
        titleRow.setHeight((short) 400);

        // Создаем изображение диаграммы
        ByteArrayOutputStream chartImage = new ByteArrayOutputStream();
        org.jfree.chart.ChartUtils.writeChartAsPNG(chartImage, chart, CHART_WIDTH, CHART_HEIGHT);

        // Добавляем изображение в книгу
        int pictureIdx = workbook.addPicture(chartImage.toByteArray(), Workbook.PICTURE_TYPE_PNG);

        // Вставляем изображение в лист
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(0);
        anchor.setRow1(startRow + 1);
        anchor.setCol2(8); // Ширина изображения в колонках
        anchor.setRow2(startRow + 25); // Высота изображения в строках

        drawing.createPicture(anchor, pictureIdx);

        return startRow + 26; // Возвращаем следующую свободную строку
    }

    // Вспомогательные методы для создания листов

    private void createStudentResultsSheet(XSSFWorkbook workbook, TestSummaryDto test,
                                           List<StudentDetailedResultDto> students) {
        XSSFSheet sheet = workbook.createSheet("Результаты студентов");

        // Определяем максимальное количество заданий
        int maxTasks = students.stream()
                .map(StudentDetailedResultDto::getTaskScores)
                .filter(Objects::nonNull)
                .mapToInt(Map::size)
                .max()
                .orElse(0);

        // Создаем заголовки
        List<String> headers = new ArrayList<>(Arrays.asList(
                "№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"
        ));

        for (int i = 1; i <= maxTasks; i++) {
            headers.add("Зад. " + i);
        }

        createHeaderRow(sheet, workbook, headers.toArray(new String[0]));

        // Заполняем данные
        int rowNum = 1;
        for (int i = 0; i < students.size(); i++) {
            StudentDetailedResultDto student = students.get(i);
            XSSFRow row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(i + 1);
            setCellValue(row, 1, student.getFio());
            setCellValue(row, 2, student.getPresence());
            setCellValue(row, 3, student.getVariant());
            setCellValue(row, 4, student.getTotalScore());

            Cell percentCell = row.createCell(5);
            percentCell.setCellValue(student.getPercentageScore() != null ?
                    student.getPercentageScore() / 100.0 : 0.0);
            percentCell.setCellStyle(createPercentStyle(workbook));

            // Баллы по заданиям
            Map<Integer, Integer> taskScores = student.getTaskScores();
            for (int taskNum = 1; taskNum <= maxTasks; taskNum++) {
                Integer score = taskScores != null ? taskScores.get(taskNum) : null;
                row.createCell(5 + taskNum).setCellValue(score != null ? score : 0);
            }
        }

        autoSizeColumns(sheet, headers.size());
    }

    /**
     * Универсальный метод установки значения в ячейку
     */
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
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    private void createTaskAnalysisSheet(XSSFWorkbook workbook, TestSummaryDto test,
                                         Map<Integer, TaskStatisticsDto> statistics) {
        XSSFSheet sheet = workbook.createSheet("Анализ по заданиям");

        String[] headers = {
                "Задание", "Макс. балл", "Полностью", "Частично", "Не справилось",
                "% выполнения", "Распределение баллов"
        };

        createHeaderRow(sheet, workbook, headers);

        // Заполняем данные
        int rowNum = 1;
        for (TaskStatisticsDto stats : statistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList())) {
            XSSFRow row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(stats.getTaskNumber());
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

        autoSizeColumns(sheet, headers.length);
    }

    private void createStatisticsSheet(XSSFWorkbook workbook, TestSummaryDto test,
                                       List<StudentDetailedResultDto> students,
                                       Map<Integer, TaskStatisticsDto> taskStatistics) {
        XSSFSheet sheet = workbook.createSheet("Общая статистика");

        XSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue(
                String.format("Статистика теста: %s - %s (%s)",
                        test.getSubject(),
                        test.getClassName(),
                        test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE))
        );

        String[][] statsData = {
                {"Предмет", test.getSubject()},
                {"Класс", test.getClassName()},
                {"Дата", test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE)},
                {"Учитель", test.getTeacher()},
                {"Всего студентов", String.valueOf(test.getClassSize())},
                {"Присутствовало", String.valueOf(test.getStudentsPresent())},
                {"Отсутствовало", String.valueOf(test.getStudentsAbsent())},
                {"% присутствия", String.format("%.1f%%", test.getAttendancePercentage())},
                {"Средний балл", String.format("%.2f", test.getAverageScore())},
                {"% выполнения", String.format("%.1f%%", test.getSuccessPercentage())},
                {"Кол-во заданий", String.valueOf(taskStatistics.size())}
        };

        int rowNum = 2;
        for (String[] rowData : statsData) {
            XSSFRow row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(rowData[0]);
            row.createCell(1).setCellValue(rowData[1]);
        }

        autoSizeColumns(sheet, 2);
    }

    private void createTeacherTestSheet(XSSFWorkbook workbook, TestSummaryDto test,
                                        String teacherName, int sheetIndex) {
        // Генерируем имя листа
        String sheetName = String.format("%s_%s",
                test.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", "").substring(0,
                        Math.min(15, test.getSubject().length())),
                test.getClassName());

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        // Проверяем уникальность имени
        String finalSheetName = sheetName;
        int counter = 1;
        while (workbook.getSheet(finalSheetName) != null) {
            finalSheetName = String.format("%s_%d", sheetName, counter++);
            if (finalSheetName.length() > 31) {
                finalSheetName = finalSheetName.substring(0, 31);
            }
        }

        // Создаем пустой лист - детальный отчет будет создан отдельно
        workbook.createSheet(finalSheetName);
    }

    /**
     * Находит последнюю заполненную строку
     */
    private int findLastRow(Sheet sheet) {
        int lastRow = sheet.getLastRowNum();
        while (lastRow >= 0 && sheet.getRow(lastRow) == null) {
            lastRow--;
        }
        return lastRow;
    }

    /**
     * Создает секцию с результатами студентов
     */
    private void createStudentResultsSection(Sheet sheet,
                                             List<StudentDetailedResultDto> students,
                                             int startRow) {
        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow);
        sectionHeader.createCell(0).setCellValue("Результаты студентов:");
        CellStyle headerStyle = createHeaderStyle(sheet.getWorkbook());
        sectionHeader.getCell(0).setCellStyle(headerStyle);

        // Определяем максимальное количество заданий
        int maxTasks = students.stream()
                .map(StudentDetailedResultDto::getTaskScores)
                .filter(Objects::nonNull)
                .mapToInt(Map::size)
                .max()
                .orElse(0);

        // Создаем заголовки таблицы
        int headerRowNum = startRow + 1;
        Row headerRow = sheet.createRow(headerRowNum);

        List<String> headers = new ArrayList<>(Arrays.asList(
                "№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"
        ));

        for (int i = 1; i <= maxTasks; i++) {
            headers.add("Зад. " + i);
        }

        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(createHeaderStyle(sheet.getWorkbook()));
        }

        // Заполняем данные
        int rowNum = headerRowNum + 1;
        for (int i = 0; i < students.size(); i++) {
            StudentDetailedResultDto student = students.get(i);
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(i + 1);
            setCellValue(row, 1, student.getFio());
            setCellValue(row, 2, student.getPresence());
            setCellValue(row, 3, student.getVariant());
            setCellValue(row, 4, student.getTotalScore());

            Cell percentCell = row.createCell(5);
            double percentage = student.getPercentageScore() != null ?
                    student.getPercentageScore() / 100.0 : 0.0;
            percentCell.setCellValue(percentage);
            percentCell.setCellStyle(createPercentStyle(sheet.getWorkbook()));

            // Баллы по заданиям
            Map<Integer, Integer> taskScores = student.getTaskScores();
            for (int taskNum = 1; taskNum <= maxTasks; taskNum++) {
                Integer score = taskScores != null ? taskScores.get(taskNum) : null;
                row.createCell(5 + taskNum).setCellValue(score != null ? score : 0);
            }
        }
    }

    /**
     * Создает секцию с анализом заданий
     */
    private void createTaskAnalysisSection(Sheet sheet,
                                           Map<Integer, TaskStatisticsDto> statistics,
                                           int startRow) {
        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow);
        sectionHeader.createCell(0).setCellValue("Анализ по заданиям:");
        CellStyle headerStyle = createHeaderStyle(sheet.getWorkbook());
        sectionHeader.getCell(0).setCellStyle(headerStyle);

        // Заголовки таблицы
        int headerRowNum = startRow + 1;
        Row headerRow = sheet.createRow(headerRowNum);

        String[] headers = {
                "Задание", "Макс. балл", "Полностью", "Частично", "Не справилось",
                "% выполнения", "Распределение баллов"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createHeaderStyle(sheet.getWorkbook()));
        }

        // Заполняем данные
        int rowNum = headerRowNum + 1;
        for (TaskStatisticsDto stats : statistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList())) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(stats.getTaskNumber());
            row.createCell(1).setCellValue(stats.getMaxScore());
            row.createCell(2).setCellValue(stats.getFullyCompletedCount());
            row.createCell(3).setCellValue(stats.getPartiallyCompletedCount());
            row.createCell(4).setCellValue(stats.getNotCompletedCount());

            Cell percentCell = row.createCell(5);
            percentCell.setCellValue(stats.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(createPercentStyle(sheet.getWorkbook()));

            if (stats.getScoreDistribution() != null) {
                String distribution = stats.getScoreDistribution().entrySet().stream()
                        .sorted(Map.Entry.comparingByKey())
                        .map(e -> String.format("%d баллов: %d", e.getKey(), e.getValue()))
                        .collect(Collectors.joining("; "));
                row.createCell(6).setCellValue(distribution);
            }
        }
    }

    // Вспомогательные методы

    private void createSummarySheetHeader(XSSFSheet sheet, XSSFWorkbook workbook) {
        String[] headers = {
                "Школа", "Предмет", "Класс", "Дата теста", "Тип теста", "Учитель",
                "Присутствовало", "Отсутствовало", "Всего в классе", "% присутствия",
                "Кол-во заданий теста", "Макс. балл", "Средний балл теста", "Файл"
        };

        XSSFRow headerRow = sheet.createRow(0);
        CellStyle headerStyle = createHeaderStyle(workbook);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void fillSummaryData(XSSFSheet sheet, XSSFWorkbook workbook, List<TestSummaryDto> tests) {
        int rowNum = 1;
        for (TestSummaryDto test : tests) {
            XSSFRow row = sheet.createRow(rowNum++);

            // Школа
            setCellValue(row, 0, test.getSchool() != null ? test.getSchool() : "ГБОУ №7");

            // Предмет
            setCellValue(row, 1, test.getSubject());

            // Класс
            setCellValue(row, 2, test.getClassName());

            // Дата теста
            Cell dateCell = row.createCell(3);
            if (test.getTestDate() != null) {
                dateCell.setCellValue(test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            }
            dateCell.setCellStyle(createCenteredStyle(workbook));

            // Тип теста
            setCellValue(row, 4, test.getTestType());

            // Учитель
            setCellValue(row, 5, test.getTeacher());

            // Числовые данные
            setNumericCellValue(row, 6, test.getStudentsPresent());
            setNumericCellValue(row, 7, test.getStudentsAbsent());
            setNumericCellValue(row, 8, test.getClassSize());

            // Процент присутствия
            Cell percentCell = row.createCell(9);
            percentCell.setCellValue(test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0);
            percentCell.setCellStyle(createPercentStyle(workbook));

            // Остальные числовые данные
            setNumericCellValue(row, 10, test.getTaskCount());
            setNumericCellValue(row, 11, test.getMaxTotalScore());

            // Средний балл
            Cell avgScoreCell = row.createCell(12);
            avgScoreCell.setCellValue(test.getAverageScore() != null ? test.getAverageScore() : 0.0);
            avgScoreCell.setCellStyle(createDecimalStyle(workbook));

            // Файл
            setCellValue(row, 13, test.getFileName());
        }
    }

    private void addSummaryStatisticsRow(XSSFSheet sheet, List<TestSummaryDto> tests) {
        if (tests.isEmpty()) return;

        int lastRow = sheet.getLastRowNum() + 1;
        XSSFRow summaryRow = sheet.createRow(lastRow);
        CellStyle summaryStyle = createSummaryStyle((XSSFWorkbook) sheet.getWorkbook());

        summaryRow.createCell(0).setCellValue("Средние показатели:");
        summaryRow.getCell(0).setCellStyle(summaryStyle);

        // Рассчитываем средние значения
        double[] averages = calculateAverages(tests);

        for (int i = 1; i < 14; i++) {
            Cell cell = summaryRow.createCell(i);
            cell.setCellStyle(summaryStyle);

            if (i >= 6 && i <= 12) {
                cell.setCellValue(averages[i - 6]);
            }
        }
    }

    private double[] calculateAverages(List<TestSummaryDto> tests) {
        return new double[]{
                tests.stream().filter(t -> t.getStudentsPresent() != null)
                        .mapToInt(TestSummaryDto::getStudentsPresent).average().orElse(0),
                tests.stream().filter(t -> t.getStudentsAbsent() != null)
                        .mapToInt(TestSummaryDto::getStudentsAbsent).average().orElse(0),
                tests.stream().filter(t -> t.getClassSize() != null)
                        .mapToInt(TestSummaryDto::getClassSize).average().orElse(0),
                tests.stream().filter(t -> t.getAttendancePercentage() != null)
                        .mapToDouble(t -> t.getAttendancePercentage() / 100.0).average().orElse(0),
                tests.stream().filter(t -> t.getTaskCount() != null)
                        .mapToInt(TestSummaryDto::getTaskCount).average().orElse(0),
                tests.stream().filter(t -> t.getMaxTotalScore() != null)
                        .mapToInt(TestSummaryDto::getMaxTotalScore).average().orElse(0),
                tests.stream().filter(t -> t.getAverageScore() != null)
                        .mapToDouble(TestSummaryDto::getAverageScore).average().orElse(0)
        };
    }

    private void createHeaderRow(Sheet sheet, XSSFWorkbook workbook, String[] headers) {
        createHeaderRow(sheet, workbook, headers, 0);
    }

    private void createHeaderRow(Sheet sheet, XSSFWorkbook workbook, String[] headers, int startRow) {
        Row headerRow = sheet.createRow(startRow);
        CellStyle headerStyle = createHeaderStyle(workbook);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void setCellValue(Row row, int cellIndex, String value) {
        if (value != null) {
            row.createCell(cellIndex).setCellValue(value);
        } else {
            row.createCell(cellIndex).setCellValue("");
        }
    }

    private void setNumericCellValue(Row row, int cellIndex, Integer value) {
        Cell cell = row.createCell(cellIndex);
        if (value != null) {
            cell.setCellValue(value);
        } else {
            cell.setCellValue(0);
        }
        cell.setCellStyle(createNumberStyle(row.getSheet().getWorkbook()));
    }

    private double calculateAverageScore(TaskStatisticsDto task) {
        if (task.getScoreDistribution() == null || task.getScoreDistribution().isEmpty()) {
            return 0.0;
        }

        double totalScore = 0;
        int totalStudents = 0;

        for (Map.Entry<Integer, Integer> entry : task.getScoreDistribution().entrySet()) {
            totalScore += entry.getKey() * entry.getValue();
            totalStudents += entry.getValue();
        }

        return totalStudents > 0 ? totalScore / totalStudents : 0.0;
    }

    private int calculateDifficultyLevel(double completionPercentage) {
        if (completionPercentage >= 90) return 1;
        else if (completionPercentage >= 70) return 2;
        else if (completionPercentage >= 50) return 3;
        else if (completionPercentage >= 30) return 4;
        else return 5;
    }

    private String getDifficultyName(int level) {
        switch (level) {
            case 1:
                return "Очень легко";
            case 2:
                return "Легко";
            case 3:
                return "Средне";
            case 4:
                return "Сложно";
            case 5:
                return "Очень сложно";
            default:
                return "Не определено";
        }
    }

    private Path createReportsFolder() throws IOException {
        Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
        if (!Files.exists(reportsPath)) {
            Files.createDirectories(reportsPath);
        }
        return reportsPath;
    }

    private File saveWorkbook(XSSFWorkbook workbook, Path reportsPath, String fileName) throws IOException {
        Path filePath = reportsPath.resolve(fileName);

        try (FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {
            workbook.write(outputStream);
        }

        workbook.close();
        log.info("Файл сохранен: {}", filePath);
        return filePath.toFile();
    }

    private void autoSizeColumns(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Создает лист с графиками с использованием JFreeChart
     */
    private void createChartsSheet(XSSFWorkbook workbook, TestSummaryDto testSummary,
                                   Map<Integer, TaskStatisticsDto> taskStatistics) {
        XSSFSheet sheet = workbook.createSheet("Графики анализа");

        // Заголовок листа с датой генерации
        XSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue(
                String.format("Графический анализ теста: %s - %s (%s)",
                        testSummary.getSubject(),
                        testSummary.getClassName(),
                        testSummary.getTestDate().format(DateTimeFormatters.DISPLAY_DATE))
        );

        XSSFRow subtitleRow = sheet.createRow(1);
        subtitleRow.createCell(0).setCellValue(
                String.format("Отчет сгенерирован: %s",
                        LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE))
        );

        int currentRow = 3; // Начинаем с четвертой строки

        try {
            // 1. Исправленная гистограмма с накоплением
            currentRow = insertChartAsImage(workbook, sheet,
                    createEnhancedStackedBarChart(taskStatistics),
                    "Распределение результатов по заданиям",
                    currentRow);

            // 2. Исправленная горизонтальная гистограмма
            currentRow = insertChartAsImage(workbook, sheet,
                    createFixedHorizontalBarChart(taskStatistics),
                    "Процент выполнения заданий",
                    currentRow + 2);

            // 3. Линейный график с тенденцией
            currentRow = insertChartAsImage(workbook, sheet,
                    createTrendLineChart(taskStatistics),
                    "Динамика выполнения заданий с линией тренда",
                    currentRow + 2);

        } catch (Exception e) {
            log.error("Ошибка при создании графиков", e);
            XSSFRow errorRow = sheet.createRow(currentRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании графиков: " + e.getMessage());
        }

        // Настройка ширины колонок
        sheet.setColumnWidth(0, IMAGE_COL_WIDTH * 256);
    }

    /**
     * Добавляет пояснения к графикам
     */
    private void addChartExplanations(XSSFSheet sheet, int startRow) {
        Row explanationsTitle = sheet.createRow(startRow++);
        explanationsTitle.createCell(0).setCellValue("Пояснения к графикам:");

        String[][] explanations = {
                {"📊 Распределение результатов по заданиям",
                        "Показывает сколько студентов полностью/частично/не справились с каждым заданием"},
                {"📈 Процент выполнения заданий",
                        "Наглядное сравнение процента выполнения всех заданий"},
                {"📉 Динамика выполнения",
                        "Показывает изменение сложности заданий по тесту. Красная линия - тренд"},
                {"🎨 Тепловая карта",
                        "Интенсивность цвета показывает сложность задания (чем темнее - тем сложнее)"},
                {"📦 Распределение баллов",
                        "Показывает разброс баллов: медиана (линия), 25-75% (ящик), min-max (усы)"}
        };

        for (String[] exp : explanations) {
            Row row = sheet.createRow(startRow++);
            row.createCell(0).setCellValue(exp[0]);
            row.createCell(1).setCellValue(exp[1]);
        }
    }

    /**
     * Создает улучшенную гистограмму с накоплением
     */
    private JFreeChart createEnhancedStackedBarChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (TaskStatisticsDto task : sortedTasks) {
            // Используем просто номер задания как метку
            String taskLabel = String.valueOf(task.getTaskNumber());

            // Добавляем абсолютные значения
            dataset.addValue(task.getFullyCompletedCount(), "✓ Полностью", taskLabel);
            dataset.addValue(task.getPartiallyCompletedCount(), "~ Частично", taskLabel);
            dataset.addValue(task.getNotCompletedCount(), "✗ Не справилось", taskLabel);
        }

        JFreeChart chart = ChartFactory.createStackedBarChart(
                "",
                "Номер задания",
                "Количество студентов",
                dataset,
                PlotOrientation.VERTICAL,
                true,
                true,
                false
        );

        // Современный дизайн
        chart.setBackgroundPaint(new java.awt.Color(245, 245, 245));

        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(java.awt.Color.WHITE);
        plot.setRangeGridlinePaint(new java.awt.Color(200, 200, 200));
        plot.setDomainGridlinePaint(new java.awt.Color(200, 200, 200));
        plot.setOutlineVisible(false);

        // Настраиваем ось X для правильного отображения номеров заданий
        CategoryAxis domainAxis = plot.getDomainAxis();
        domainAxis.setTickLabelFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 10));
        domainAxis.setCategoryMargin(0.1);

        // Настраиваем ось Y
        ValueAxis rangeAxis = plot.getRangeAxis();
        rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());

        // Современная цветовая палитра
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setSeriesPaint(0, new java.awt.Color(46, 204, 113));   // Ярко-зеленый
        renderer.setSeriesPaint(1, new java.awt.Color(241, 196, 15));   // Ярко-желтый
        renderer.setSeriesPaint(2, new java.awt.Color(231, 76, 60));    // Ярко-красный

        // Используем наш ModernGradientBarPainter
        renderer.setBarPainter(new ModernGradientBarPainter());
        renderer.setShadowVisible(false);

        // Добавляем значения на столбцы
        renderer.setDefaultItemLabelGenerator(new CategoryItemLabelGenerator() {
            @Override
            public String generateRowLabel(org.jfree.data.category.CategoryDataset dataset, int row) {
                return "";
            }

            @Override
            public String generateColumnLabel(org.jfree.data.category.CategoryDataset dataset, int column) {
                return "";
            }

            @Override
            public String generateLabel(org.jfree.data.category.CategoryDataset dataset, int row, int column) {
                Number value = dataset.getValue(row, column);
                return value.intValue() > 0 ? String.valueOf(value.intValue()) : "";
            }
        });

        renderer.setDefaultItemLabelsVisible(true);

        // Шрифт для подписей значений
        java.awt.Font labelFont = new java.awt.Font("Arial", java.awt.Font.PLAIN, 9);
        renderer.setDefaultItemLabelFont(labelFont);

        // Позиционируем подписи внутри столбцов
        renderer.setDefaultPositiveItemLabelPosition(new ItemLabelPosition(
                ItemLabelAnchor.CENTER, TextAnchor.CENTER, TextAnchor.CENTER, 0.0));

        // Легенда
        org.jfree.chart.title.LegendTitle legend = chart.getLegend();
        legend.setBackgroundPaint(java.awt.Color.WHITE);
        legend.setFrame(new org.jfree.chart.block.BlockBorder(java.awt.Color.LIGHT_GRAY));

        // Шрифт для легенды
        java.awt.Font legendFont = new java.awt.Font("Arial", java.awt.Font.PLAIN, 10);
        legend.setItemFont(legendFont);

        return chart;
    }
    /**
     * Создает горизонтальную гистограмму с правильным позиционированием процентов
     */
    private JFreeChart createFixedHorizontalBarChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (TaskStatisticsDto task : sortedTasks) {
            String taskLabel = "Задание " + task.getTaskNumber();
            dataset.addValue(task.getCompletionPercentage(), "% выполнения", taskLabel);
        }

        JFreeChart chart = ChartFactory.createBarChart(
                "",
                "Процент выполнения, %",
                "Задание",
                dataset,
                PlotOrientation.HORIZONTAL,
                false, // без легенды
                true,
                false
        );

        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(java.awt.Color.WHITE);
        plot.setRangeGridlinePaint(new java.awt.Color(200, 200, 200));
        plot.setDomainGridlinePaint(new java.awt.Color(200, 200, 200));

        // Настраиваем оси
        CategoryAxis domainAxis = plot.getDomainAxis();
        domainAxis.setTickLabelFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 10));

        NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
        rangeAxis.setRange(0, 100); // Проценты от 0 до 100
        rangeAxis.setStandardTickUnits(NumberAxis.createStandardTickUnits());
        rangeAxis.setNumberFormatOverride(java.text.NumberFormat.getPercentInstance());

        // Градиентная заливка
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBarPainter(new ModernGradientBarPainter());
        renderer.setShadowVisible(false);

        // Цветовая шкала от красного к зеленому в зависимости от процента
        for (int i = 0; i < sortedTasks.size(); i++) {
            double percentage = sortedTasks.get(i).getCompletionPercentage();
            java.awt.Color color = getColorForPercentage(percentage);
            renderer.setSeriesPaint(i, color);
        }

        // Добавляем значения ПРАВИЛЬНО позиционированные
        renderer.setDefaultItemLabelGenerator(new CategoryItemLabelGenerator() {
            @Override
            public String generateRowLabel(org.jfree.data.category.CategoryDataset dataset, int row) {
                return "";
            }

            @Override
            public String generateColumnLabel(org.jfree.data.category.CategoryDataset dataset, int column) {
                return "";
            }

            @Override
            public String generateLabel(org.jfree.data.category.CategoryDataset dataset, int row, int column) {
                Number value = dataset.getValue(row, column);
                return String.format("%.1f%%", value.doubleValue());
            }
        });

        renderer.setDefaultItemLabelsVisible(true);

        // Шрифт для процентов - УВЕЛИЧИВАЕМ и делаем ПОЛУЖИРНЫМ
        java.awt.Font percentFont = new java.awt.Font("Arial", java.awt.Font.BOLD, 10);
        renderer.setDefaultItemLabelFont(percentFont);

        // Ключевое исправление: Позиционируем проценты ВНУТРИ столбцов, справа от начала
        renderer.setDefaultPositiveItemLabelPosition(new ItemLabelPosition(
                ItemLabelAnchor.OUTSIDE3, // Внутри столбца, слева от значения
                TextAnchor.CENTER_LEFT,
                TextAnchor.CENTER_LEFT,
                Math.PI / 2.0 // Поворачиваем на 0 градусов (горизонтально)
        ));

        // Устанавливаем отступ для подписей, чтобы они были внутри столбцов
        renderer.setItemLabelAnchorOffset(10.0);

        return chart;
    }

    /**
     * Создает горизонтальную гистограмму с процентами
     */
    private JFreeChart createHorizontalBarChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (TaskStatisticsDto task : sortedTasks) {
            String taskLabel = "Задание " + task.getTaskNumber();
            dataset.addValue(task.getCompletionPercentage(), "% выполнения", taskLabel);
        }

        JFreeChart chart = ChartFactory.createBarChart(
                "",
                "Процент выполнения, %",
                "Задание",
                dataset,
                PlotOrientation.HORIZONTAL,
                false, // без легенды
                true,
                false
        );

        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setRangeGridlinePaint(new Color(200, 200, 200));

        // Градиентная заливка в зависимости от процента
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBarPainter(new GradientBarPainter(0.0, 0.0, 0.0));
        renderer.setShadowVisible(false);

        // Цветовая шкала от красного к зеленому
        for (int i = 0; i < sortedTasks.size(); i++) {
            double percentage = sortedTasks.get(i).getCompletionPercentage();
            Color color = getColorForPercentage(percentage);
            renderer.setSeriesPaint(i, color);
        }

        // Добавляем значения в концы столбцов
        renderer.setDefaultItemLabelGenerator(new CategoryItemLabelGenerator() {
            @Override
            public String generateRowLabel(CategoryDataset dataset, int row) {
                return "";
            }

            @Override
            public String generateColumnLabel(CategoryDataset dataset, int column) {
                return "";
            }

            @Override
            public String generateLabel(CategoryDataset dataset, int row, int column) {
                Number value = dataset.getValue(row, column);
                return String.format("%.1f%%", value.doubleValue());
            }
        });
        renderer.setDefaultItemLabelsVisible(true);
        java.awt.Font labelFont = new java.awt.Font("Arial", java.awt.Font.BOLD, 10);
        renderer.setDefaultItemLabelFont(labelFont);
        renderer.setDefaultPositiveItemLabelPosition(new ItemLabelPosition(
                ItemLabelAnchor.OUTSIDE12, TextAnchor.CENTER_LEFT, TextAnchor.CENTER_LEFT, 0.0));

        return chart;
    }

    /**
     * Создает тепловую карту выполнения заданий
     */
    private JFreeChart createHeatMapChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (TaskStatisticsDto task : sortedTasks) {
            dataset.addValue(task.getCompletionPercentage(), "Выполнение",
                    String.valueOf(task.getTaskNumber()));
        }

        JFreeChart chart = ChartFactory.createBarChart(
                "",
                "Номер задания",
                "% выполнения",
                dataset,
                PlotOrientation.VERTICAL,
                false,
                true,
                false
        );

        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);

        // Используем градиентные цвета
        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setBarPainter(new GradientBarPainter(0.0, 0.0, 0.0));

        // Настраиваем цвета в зависимости от процента выполнения
        for (int i = 0; i < sortedTasks.size(); i++) {
            double percentage = sortedTasks.get(i).getCompletionPercentage();
            renderer.setSeriesPaint(i, getHeatMapColor(percentage));
        }

        return chart;
    }

    /**
     * Создает box plot для распределения баллов
     */
    private JFreeChart createBoxPlotChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultBoxAndWhiskerCategoryDataset dataset = new DefaultBoxAndWhiskerCategoryDataset();

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (TaskStatisticsDto task : sortedTasks) {
            List<Double> scores = new ArrayList<>();

            if (task.getScoreDistribution() != null) {
                task.getScoreDistribution().forEach((score, count) -> {
                    for (int i = 0; i < count; i++) {
                        scores.add(score.doubleValue());
                    }
                });
            }

            if (!scores.isEmpty()) {
                dataset.add(scores, "Задание", String.valueOf(task.getTaskNumber()));
            }
        }

        JFreeChart chart = ChartFactory.createBoxAndWhiskerChart(
                "",
                "Номер задания",
                "Баллы",
                dataset,
                true
        );

        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setDomainGridlinePaint(Color.LIGHT_GRAY);
        plot.setRangeGridlinePaint(Color.LIGHT_GRAY);

        BoxAndWhiskerRenderer renderer = (BoxAndWhiskerRenderer) plot.getRenderer();
        renderer.setFillBox(true);
        renderer.setSeriesPaint(0, new Color(52, 152, 219, 180));
        renderer.setMeanVisible(true);
        renderer.setMedianVisible(true);

        return chart;
    }

    /**
     * Возвращает цвет для процента выполнения (от красного к зеленому)
     */
    private Color getColorForPercentage(double percentage) {
        if (percentage >= 90) return new Color(39, 174, 96);    // Зеленый
        if (percentage >= 70) return new Color(46, 204, 113);   // Светло-зеленый
        if (percentage >= 50) return new Color(241, 196, 15);   // Желтый
        if (percentage >= 30) return new Color(230, 126, 34);   // Оранжевый
        return new Color(231, 76, 60);                         // Красный
    }

    /**
     * Создает линейный график с линией тренда
     */
    private JFreeChart createTrendLineChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        XYSeries completionSeries = new XYSeries("% выполнения");
        XYSeries avgScoreSeries = new XYSeries("Средний балл");
        XYSeries trendSeries = new XYSeries("Тренд");

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        List<Double> completionValues = new ArrayList<>();

        for (int i = 0; i < sortedTasks.size(); i++) {
            TaskStatisticsDto task = sortedTasks.get(i);
            double completion = task.getCompletionPercentage();
            completionSeries.add(i + 1, completion);
            completionValues.add(completion);

            double avgScore = calculateAverageScore(task);
            double normalizedScore = (avgScore / task.getMaxScore()) * 100;
            avgScoreSeries.add(i + 1, normalizedScore);
        }

        // Расчет линейного тренда
        if (completionValues.size() > 1) {
            double[] trend = calculateLinearTrend(completionValues);
            for (int i = 0; i < sortedTasks.size(); i++) {
                trendSeries.add(i + 1, trend[0] * (i + 1) + trend[1]);
            }
        }

        XYSeriesCollection dataset = new XYSeriesCollection();
        dataset.addSeries(completionSeries);
        dataset.addSeries(avgScoreSeries);
        dataset.addSeries(trendSeries);

        JFreeChart chart = ChartFactory.createXYLineChart(
                "",
                "Номер задания",
                "Показатель, %",
                dataset,
                PlotOrientation.VERTICAL,
                true,
                true,
                false
        );

        XYPlot plot = chart.getXYPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setDomainGridlinePaint(new Color(220, 220, 220));
        plot.setRangeGridlinePaint(new Color(220, 220, 220));

        // Современные стили линий
        XYLineAndShapeRenderer renderer = new XYLineAndShapeRenderer();

        // Основная линия - % выполнения
        renderer.setSeriesPaint(0, new Color(52, 152, 219));
        renderer.setSeriesStroke(0, new BasicStroke(2.5f));
        renderer.setSeriesShape(0,
                new java.awt.geom.Ellipse2D.Double(-4, -4, 8, 8));

        // Вторая линия - средний балл
        renderer.setSeriesPaint(1, new Color(155, 89, 182));
        renderer.setSeriesStroke(1, new BasicStroke(2.0f, BasicStroke.CAP_ROUND,
                BasicStroke.JOIN_ROUND, 1.0f, new float[]{6.0f, 6.0f}, 0.0f));
        renderer.setSeriesShape(1,
                new java.awt.Rectangle(-3, -3, 6, 6));

        // Линия тренда
        renderer.setSeriesPaint(2, new Color(231, 76, 60));
        renderer.setSeriesStroke(2, new BasicStroke(1.5f, BasicStroke.CAP_BUTT,
                BasicStroke.JOIN_BEVEL, 0.0f, new float[]{10.0f, 6.0f}, 0.0f));
        renderer.setSeriesShapesVisible(2, false);

        plot.setRenderer(renderer);

        // Настройка легенды
        chart.getLegend().setBackgroundPaint(Color.WHITE);
        chart.getLegend().setFrame(new BlockBorder(Color.LIGHT_GRAY));

        return chart;
    }

    /**
     * Расчет линейного тренда методом наименьших квадратов
     */
    private double[] calculateLinearTrend(List<Double> values) {
        int n = values.size();
        double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

        for (int i = 0; i < n; i++) {
            double x = i + 1;
            double y = values.get(i);
            sumX += x;
            sumY += y;
            sumXY += x * y;
            sumX2 += x * x;
        }

        double slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        double intercept = (sumY - slope * sumX) / n;

        return new double[]{slope, intercept};
    }

    /**
     * Возвращает цвет для тепловой карты
     */
    private Color getHeatMapColor(double percentage) {
        // От темно-синего (сложное) к светло-зеленому (легкое)
        if (percentage >= 90) return new Color(162, 222, 150);  // Светло-зеленый
        if (percentage >= 70) return new Color(117, 199, 111);  // Зеленый
        if (percentage >= 50) return new Color(72, 176, 72);    // Темно-зеленый
        if (percentage >= 30) return new Color(44, 127, 184);   // Синий
        return new Color(37, 52, 148);                         // Темно-синий
    }

    // Стили

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setBorders(style);
        return style;
    }

    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        setBorders(style);
        return style;
    }

    private CellStyle createCenteredStyle(Workbook workbook) {
        CellStyle style = createDataStyle(workbook);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createNumberStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0"));
        setBorders(style);
        return style;
    }

    private CellStyle createDecimalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        setBorders(style);
        return style;
    }

    private CellStyle createPercentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        setBorders(style);
        return style;
    }

    private CellStyle createSummaryStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        setBorders(style);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        return style;
    }

    private void setBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    /**
     * Создает стиль для основного заголовка
     */
    private CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFont(font);
        return style;
    }

    /**
     * Создает стиль для подзаголовка
     */
    private CellStyle createSubtitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setItalic(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.DARK_GREEN.getIndex());
        style.setFont(font);
        return style;
    }

    /**
     * Создает стиль для заголовков секций
     */
    private CellStyle createSectionHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.ROYAL_BLUE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBottomBorderColor(IndexedColors.DARK_BLUE.getIndex());
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
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        setBorders(style);
        return style;
    }
}