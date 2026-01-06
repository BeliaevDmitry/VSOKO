package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.plot.*;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
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

        log.info("Генерация детального отчета для теста: {} - {}",
                testSummary.getSubject(), testSummary.getClassName());

        try {
            Path reportsPath = createReportsFolder();

            String fileName = String.format("Детальный_отчет_%s_%s_%s.xlsx",
                    testSummary.getSubject(),
                    testSummary.getClassName(),
                    testSummary.getTestDate().format(DateTimeFormatter.ofPattern("ddMMyyyy")));

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                // Основные листы
                createStudentResultsSheet(workbook, testSummary, studentResults);
                createTaskAnalysisSheet(workbook, testSummary, taskStatistics);
                createStatisticsSheet(workbook, testSummary, studentResults, taskStatistics);

                // Лист с графиками (используем JFreeChart)
                createChartsSheet(workbook, testSummary, taskStatistics);

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

    /**
     * Создает лист с графиками с использованием JFreeChart
     */
    private void createChartsSheet(XSSFWorkbook workbook, TestSummaryDto testSummary,
                                   Map<Integer, TaskStatisticsDto> taskStatistics) {
        XSSFSheet sheet = workbook.createSheet("Графики анализа");

        // Заголовок листа
        XSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue(
                String.format("Графический анализ теста: %s - %s (%s)",
                        testSummary.getSubject(),
                        testSummary.getClassName(),
                        testSummary.getTestDate().format(DateTimeFormatters.DISPLAY_DATE))
        );

        int currentRow = 2; // Начинаем с третьей строки

        try {
            // 1. Гистограмма с накоплением
            currentRow = insertChartAsImage(workbook, sheet,
                    createStackedBarChart(taskStatistics),
                    "Распределение результатов по заданиям",
                    currentRow);

            // 2. Линейный график
            currentRow = insertChartAsImage(workbook, sheet,
                    createLineChart(taskStatistics),
                    "Динамика выполнения заданий",
                    currentRow + 2);

            // 3. Круговая диаграмма
            currentRow = insertChartAsImage(workbook, sheet,
                    createPieChart(taskStatistics),
                    "Распределение заданий по сложности",
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
        ChartUtils.writeChartAsPNG(chartImage, chart, CHART_WIDTH, CHART_HEIGHT);

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

    /**
     * Создает гистограмму с накоплением
     */
    private JFreeChart createStackedBarChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (TaskStatisticsDto task : sortedTasks) {
            String taskLabel = "Зад. " + task.getTaskNumber();
            dataset.addValue(task.getFullyCompletedCount(), "Полностью справилось", taskLabel);
            dataset.addValue(task.getPartiallyCompletedCount(), "Частично справилось", taskLabel);
            dataset.addValue(task.getNotCompletedCount(), "Не справилось", taskLabel);
        }

        JFreeChart chart = ChartFactory.createStackedBarChart(
                "", // Заголовок уже есть в Excel
                "Номер задания",
                "Количество студентов",
                dataset,
                PlotOrientation.VERTICAL,
                true,
                true,
                false
        );

        // Настройка внешнего вида
        CategoryPlot plot = chart.getCategoryPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setRangeGridlinePaint(Color.GRAY);

        BarRenderer renderer = (BarRenderer) plot.getRenderer();
        renderer.setSeriesPaint(0, new Color(76, 175, 80));   // Зеленый
        renderer.setSeriesPaint(1, new Color(255, 193, 7));   // Желтый
        renderer.setSeriesPaint(2, new Color(244, 67, 54));   // Красный
        renderer.setBarPainter(new StandardBarPainter());

        chart.getLegend().setFrame(BlockBorder.NONE);

        return chart;
    }

    /**
     * Создает линейный график
     */
    private JFreeChart createLineChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        XYSeries completionSeries = new XYSeries("% выполнения");
        XYSeries avgScoreSeries = new XYSeries("Средний балл");

        List<TaskStatisticsDto> sortedTasks = taskStatistics.values().stream()
                .sorted(Comparator.comparingInt(TaskStatisticsDto::getTaskNumber))
                .collect(Collectors.toList());

        for (int i = 0; i < sortedTasks.size(); i++) {
            TaskStatisticsDto task = sortedTasks.get(i);
            completionSeries.add(i + 1, task.getCompletionPercentage());

            double avgScore = calculateAverageScore(task);
            double normalizedScore = (avgScore / task.getMaxScore()) * 100;
            avgScoreSeries.add(i + 1, normalizedScore);
        }

        XYSeriesCollection dataset = new XYSeriesCollection();
        dataset.addSeries(completionSeries);
        dataset.addSeries(avgScoreSeries);

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
        plot.setDomainGridlinePaint(Color.LIGHT_GRAY);
        plot.setRangeGridlinePaint(Color.LIGHT_GRAY);

        XYLineAndShapeRenderer renderer = new XYLineAndShapeRenderer();
        renderer.setSeriesPaint(0, new Color(33, 150, 243));    // Синий
        renderer.setSeriesPaint(1, new Color(255, 152, 0));     // Оранжевый
        renderer.setSeriesShapesVisible(0, true);
        renderer.setSeriesShapesVisible(1, true);
        renderer.setSeriesShape(0, new java.awt.geom.Ellipse2D.Double(-3, -3, 6, 6));
        renderer.setSeriesShape(1, new java.awt.Rectangle(-3, -3, 6, 6));

        plot.setRenderer(renderer);
        chart.getLegend().setFrame(BlockBorder.NONE);

        return chart;
    }

    /**
     * Создает круговую диаграмму
     */
    private JFreeChart createPieChart(Map<Integer, TaskStatisticsDto> taskStatistics) {
        DefaultPieDataset dataset = new DefaultPieDataset();

        Map<Integer, Long> difficultyDistribution = new HashMap<>();
        for (TaskStatisticsDto task : taskStatistics.values()) {
            int difficulty = calculateDifficultyLevel(task.getCompletionPercentage());
            difficultyDistribution.put(difficulty,
                    difficultyDistribution.getOrDefault(difficulty, 0L) + 1);
        }

        for (Map.Entry<Integer, Long> entry : difficultyDistribution.entrySet()) {
            String difficultyName = getDifficultyName(entry.getKey());
            dataset.setValue(difficultyName, entry.getValue());
        }

        JFreeChart chart = ChartFactory.createPieChart(
                "",
                dataset,
                true,
                true,
                false
        );

        PiePlot plot = (PiePlot) chart.getPlot();
        plot.setBackgroundPaint(Color.WHITE);
        plot.setOutlinePaint(null);

        // Цвета для уровней сложности
        plot.setSectionPaint("Очень легко", new Color(76, 175, 80));
        plot.setSectionPaint("Легко", new Color(139, 195, 74));
        plot.setSectionPaint("Средне", new Color(255, 193, 7));
        plot.setSectionPaint("Сложно", new Color(255, 152, 0));
        plot.setSectionPaint("Очень сложно", new Color(244, 67, 54));

        plot.setLabelGenerator(null);
        chart.getLegend().setFrame(BlockBorder.NONE);

        return chart;
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

    private void createTeacherSummarySheet(XSSFWorkbook workbook, String teacherName,
                                           List<TestSummaryDto> teacherTests) {
        XSSFSheet sheet = workbook.createSheet("Сводка по тестам");

        XSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue("Отчет по тестам учителя: " + teacherName);

        String[] headers = {
                "Предмет", "Класс", "Дата", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения"
        };

        createHeaderRow(sheet, workbook, headers, 2);

        // Данные
        int rowNum = 3;
        for (TestSummaryDto test : teacherTests) {
            XSSFRow row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(test.getSubject());
            row.createCell(1).setCellValue(test.getClassName());
            row.createCell(2).setCellValue(
                    test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            row.createCell(3).setCellValue(test.getTestType());
            row.createCell(4).setCellValue(test.getStudentsPresent());
            row.createCell(5).setCellValue(test.getStudentsAbsent());
            row.createCell(6).setCellValue(test.getClassSize());

            Cell attendanceCell = row.createCell(7);
            attendanceCell.setCellValue(test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0);
            attendanceCell.setCellStyle(createPercentStyle(workbook));

            row.createCell(8).setCellValue(test.getAverageScore());

            Cell successCell = row.createCell(9);
            successCell.setCellValue(test.getSuccessPercentage() != null ?
                    test.getSuccessPercentage() / 100.0 : 0.0);
            successCell.setCellStyle(createPercentStyle(workbook));
        }

        autoSizeColumns(sheet, headers.length);
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

        XSSFSheet sheet = workbook.createSheet(finalSheetName);

        XSSFRow titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue(
                String.format("%s - %s (%s)",
                        test.getSubject(),
                        test.getClassName(),
                        test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE))
        );

        // Можно добавить дополнительную информацию
        XSSFRow infoRow = sheet.createRow(1);
        infoRow.createCell(0).setCellValue("Учитель: " + test.getTeacher());

        autoSizeColumns(sheet, 1);
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
        return new double[] {
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
            case 1: return "Очень легко";
            case 2: return "Легко";
            case 3: return "Средне";
            case 4: return "Сложно";
            case 5: return "Очень сложно";
            default: return "Не определено";
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
}