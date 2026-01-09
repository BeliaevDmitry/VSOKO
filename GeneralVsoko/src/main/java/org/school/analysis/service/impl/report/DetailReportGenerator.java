package org.school.analysis.service.impl.report;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.impl.report.charts.ExcelChartService;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Component
@RequiredArgsConstructor
@Slf4j
public class DetailReportGenerator extends ExcelReportBase {

    private final ExcelChartService excelChartService;

    // ============ ПУБЛИЧНЫЕ МЕТОДЫ ============

    /**
     * Генерирует отдельный файл с детальным отчетом
     */
    public File generateDetailReportFile(
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
                String sheetName = createSafeSheetName(testSummary);

                createDetailReportOnSheet(workbook, testSummary, studentResults, taskStatistics, sheetName);

                return saveWorkbook(workbook, reportsPath, fileName);
            }

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета", e);
            return null;
        }
    }

    /**
     * Создает детальный отчет на листе в существующей рабочей книге
     */
    public void createDetailReportOnSheet(XSSFWorkbook workbook,
                                          TestSummaryDto testSummary,
                                          List<StudentDetailedResultDto> studentResults,
                                          Map<Integer, TaskStatisticsDto> taskStatistics,
                                          String sheetName) {

        Sheet sheet = workbook.createSheet(sheetName);
        int currentRow = 0;

        try {
            // 1. Заголовок отчета
            currentRow = createReportHeader(sheet, workbook, testSummary, currentRow);

            // 2. Общая статистика теста
            currentRow = createTestStatistics(sheet, workbook, testSummary, currentRow);

            // 3. Определяем количество заданий
            int maxTasks = determineMaxTaskCount(studentResults, taskStatistics);

            // 4. Результаты студентов
            currentRow = createStudentResults(sheet, workbook, studentResults, currentRow, maxTasks);

            // 5. Анализ по заданиям
            currentRow = createTaskAnalysis(sheet, workbook, taskStatistics, currentRow, maxTasks);

            // 6. Графики (если есть данные)
            if (taskStatistics != null && !taskStatistics.isEmpty() &&
                    studentResults != null && !studentResults.isEmpty()) {

                excelChartService.createCharts(workbook, sheet, testSummary, taskStatistics, currentRow);
            }

            // 7. Оптимизируем ширину колонок
            optimizeDetailReportColumns(sheet, maxTasks);

            log.info("✅ Детальный отчет создан на листе '{}': {} строк", sheetName, currentRow);

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета", e);
            Row errorRow = sheet.createRow(currentRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании отчета: " + e.getMessage());
        }
    }

    /**
     * Создает безопасное имя листа
     */
    public String createSafeSheetName(TestSummaryDto testSummary) {
        String sheetName = String.format("%s_%s",
                testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", ""),
                testSummary.getClassName());

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        return sheetName;
    }

    /**
     * Создает уникальное имя листа для рабочей книги
     */
    public String createUniqueSheetName(XSSFWorkbook workbook, TestSummaryDto testSummary) {
        String baseName = String.format("%s_%s_%s",
                testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", "").substring(0,
                        Math.min(12, testSummary.getSubject().length())),
                testSummary.getClassName(),
                testSummary.getTestDate().format(DateTimeFormatter.ofPattern("ddMM")));

        if (baseName.length() > 31) {
            baseName = baseName.substring(0, 31);
        }

        String finalName = baseName;
        int counter = 1;
        while (workbook.getSheet(finalName) != null) {
            finalName = String.format("%s_%d", baseName, counter++);
            if (finalName.length() > 31) {
                finalName = finalName.substring(0, 31);
            }
        }

        return finalName;
    }

    // ============ ПРИВАТНЫЕ ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ============

    private int createReportHeader(Sheet sheet, Workbook workbook,
                                   TestSummaryDto testSummary, int startRow) {

        // Основной заголовок
        Row titleRow = sheet.createRow(startRow++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(
                String.format("ДЕТАЛЬНЫЙ ОТЧЕТ ПО ТЕСТУ: %s - %s",
                        testSummary.getSubject(), testSummary.getClassName())
        );
        titleCell.setCellStyle(createTitleStyle(workbook));

        // Объединяем ячейки для заголовка
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, 24));

        // Подзаголовок
        Row subtitleRow = sheet.createRow(startRow++);
        Cell subtitleCell = subtitleRow.createCell(0);
        subtitleCell.setCellValue(
                String.format("Дата проведения: %s | Тип работы: %s",
                        testSummary.getTestDate().format(DateTimeFormatters.DISPLAY_DATE),
                        testSummary.getTestType())
        );
        subtitleCell.setCellStyle(createSubtitleStyle(workbook));

        // Дополнительная информация
        sheet.createRow(startRow++).createCell(0).setCellValue("Учитель: " + testSummary.getTeacher());
        sheet.createRow(startRow++).createCell(0).setCellValue("Файл отчета: " + testSummary.getFileName());

        return addEmptyRows(sheet, startRow, 1); // Пустая строка
    }

    private int createTestStatistics(Sheet sheet, Workbook workbook,
                                     TestSummaryDto testSummary, int startRow) {

        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("ОБЩАЯ СТАТИСТИКА ТЕСТА");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Данные статистики
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

        return addEmptyRows(sheet, startRow, 1);
    }

    private int determineMaxTaskCount(List<StudentDetailedResultDto> studentResults,
                                      Map<Integer, TaskStatisticsDto> taskStatistics) {

        int maxTasks = 0;

        // Проверяем статистику
        if (taskStatistics != null && !taskStatistics.isEmpty()) {
            maxTasks = taskStatistics.keySet().stream()
                    .mapToInt(Integer::intValue)
                    .max()
                    .orElse(0);
        }

        // Проверяем студентов
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

        // Если все еще 0, используем значение по умолчанию
        if (maxTasks == 0) {
            maxTasks = 10;
        }

        // Гарантируем минимум 10 заданий для корректного отображения
        return Math.max(maxTasks, 10);
    }

    private int createStudentResults(Sheet sheet, Workbook workbook,
                                     List<StudentDetailedResultDto> studentResults,
                                     int startRow, int maxTasks) {

        if (studentResults == null || studentResults.isEmpty()) {
            return startRow;
        }

        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("РЕЗУЛЬТАТЫ СТУДЕНТОВ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Формируем заголовки
        List<String> headers = new ArrayList<>(Arrays.asList(
                "№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"
        ));

        for (int i = 1; i <= maxTasks; i++) {
            headers.add("№" + i);
        }

        // Создаем строку заголовков
        Row headerRow = sheet.createRow(startRow++);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        // Заполняем данные студентов
        for (int i = 0; i < studentResults.size(); i++) {
            StudentDetailedResultDto student = studentResults.get(i);
            Row row = sheet.createRow(startRow++);

            // Базовые данные
            row.createCell(0).setCellValue(i + 1);
            setCellValue(row, 1, student.getFio());
            setCellValue(row, 2, student.getPresence());
            setCellValue(row, 3, student.getVariant());
            setCellValue(row, 4, student.getTotalScore());

            // Процент выполнения
            Cell percentCell = row.createCell(5);
            double percentage = student.getPercentageScore() != null ?
                    student.getPercentageScore() / 100.0 : 0.0;
            percentCell.setCellValue(percentage);
            percentCell.setCellStyle(createPercentStyle(workbook));

            // Баллы за задания
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

        return addEmptyRows(sheet, startRow, 1);
    }

    private int createTaskAnalysis(Sheet sheet, Workbook workbook,
                                   Map<Integer, TaskStatisticsDto> taskStatistics,
                                   int startRow, int maxTasks) {

        if (taskStatistics == null || taskStatistics.isEmpty()) {
            return startRow;
        }

        // Заголовок секции
        Row sectionHeader = sheet.createRow(startRow++);
        sectionHeader.createCell(0).setCellValue("АНАЛИЗ ПО ЗАДАНИЯМ");
        sectionHeader.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Заголовки таблицы
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

        // Заполняем данные по заданиям
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
                // Значения по умолчанию для отсутствующих заданий
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

        return addEmptyRows(sheet, startRow, 1);
    }
}