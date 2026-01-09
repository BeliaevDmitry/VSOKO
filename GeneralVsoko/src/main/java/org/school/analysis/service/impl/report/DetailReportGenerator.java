package org.school.analysis.service.impl.report;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.impl.report.charts.ExcelChartService;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class DetailReportGenerator {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final Path REPORTS_DIR = Paths.get("reports", "detailed");

    private final ExcelChartService excelChartService;

    public File generateDetailReportFile(TestSummaryDto testSummary,
                                         List<StudentDetailedResultDto> studentResults,
                                         Map<Integer, TaskStatisticsDto> taskStatistics) {

        log.info("Генерация детального отчета для теста: {}", testSummary.getFileName());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Отчет по тесту");

            // Начинаем с первой строки
            int currentRow = 0;

            // 1. Заголовок отчета
            currentRow = createReportHeader(workbook, sheet, testSummary, currentRow);

            // 2. Общая статистика
            currentRow = createGeneralStatistics(workbook, sheet, testSummary, currentRow);

            // 3. Результаты студентов
            currentRow = createStudentsResults(workbook, sheet, studentResults, taskStatistics, currentRow);

            // 4. Анализ по заданиям (основная таблица)
            currentRow = createTaskAnalysis(workbook, sheet, taskStatistics, currentRow);

            // 5. ГРАФИЧЕСКИЙ АНАЛИЗ (ТОЛЬКО ГРАФИКИ, без дублирования данных)
            currentRow = createGraphicalAnalysis(workbook, sheet, testSummary, taskStatistics, currentRow);

            // Автонастройка ширины столбцов
            autoSizeColumns(sheet, taskStatistics.size());

            // Сохраняем файл
            return saveWorkbookToFile(workbook, testSummary);

        } catch (Exception e) {
            log.error("Ошибка генерации детального отчета: {}", e.getMessage(), e);
            return null;
        }
    }

    /**
     * Создает раздел с графиками (БЕЗ дублирования табличных данных)
     */
    private int createGraphicalAnalysis(XSSFWorkbook workbook, XSSFSheet sheet,
                                        TestSummaryDto testSummary,
                                        Map<Integer, TaskStatisticsDto> taskStatistics,
                                        int startRow) {

        log.debug("Создание графического анализа, стартовая строка: {}", startRow);

        // Оставляем пустую строку перед графиками
        startRow += 2;

        // Используем ExcelChartService для создания графиков
        excelChartService.createCharts(workbook, sheet, testSummary, taskStatistics, startRow);

        // Возвращаем следующую свободную строку
        return sheet.getLastRowNum() + 3;
    }

    /**
     * Автонастройка ширины столбцов с учетом количества заданий
     */
    private void autoSizeColumns(XSSFSheet sheet, int taskCount) {
        int columnsToAutoSize = 8 + taskCount; // Базовые столбцы + задания
        columnsToAutoSize = Math.min(columnsToAutoSize, 50); // Ограничиваем

        for (int i = 0; i < columnsToAutoSize; i++) {
            sheet.autoSizeColumn(i);
            // Устанавливаем минимальную ширину
            if (sheet.getColumnWidth(i) < 2000) {
                sheet.setColumnWidth(i, 2000);
            }
            // Ограничиваем максимальную ширину
            if (sheet.getColumnWidth(i) > 8000) {
                sheet.setColumnWidth(i, 8000);
            }
        }
    }

    // ============ ОСНОВНЫЕ МЕТОДЫ СОЗДАНИЯ ОТЧЕТА ============

    /**
     * Создает заголовок отчета
     */
    private int createReportHeader(Workbook workbook, Sheet sheet,
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
                        testSummary.getTestDate().format(DATE_FORMATTER),
                        testSummary.getTestType())
        );
        subtitleCell.setCellStyle(createSubtitleStyle(workbook));

        // Дополнительная информация
        if (testSummary.getTeacher() != null && !testSummary.getTeacher().isEmpty()) {
            sheet.createRow(startRow++).createCell(0).setCellValue("Учитель: " + testSummary.getTeacher());
        }

        if (testSummary.getFileName() != null && !testSummary.getFileName().isEmpty()) {
            sheet.createRow(startRow++).createCell(0).setCellValue("Файл отчета: " + testSummary.getFileName());
        }

        return startRow + 1; // Пустая строка
    }

    /**
     * Создает общую статистику теста
     */
    private int createGeneralStatistics(Workbook workbook, Sheet sheet,
                                        TestSummaryDto testSummary, int startRow) {

        // Заголовок раздела
        Row headerRow = sheet.createRow(startRow++);
        headerRow.createCell(0).setCellValue("ОБЩАЯ СТАТИСТИКА ТЕСТА");
        headerRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Данные статистики
        startRow = addStatisticRow(sheet, startRow, "Всего учеников в классе",
                testSummary.getClassSize() != null ? testSummary.getClassSize() : testSummary.getStudentsTotal());
        startRow = addStatisticRow(sheet, startRow, "Присутствовало на тесте", testSummary.getStudentsPresent());
        startRow = addStatisticRow(sheet, startRow, "Отсутствовало на тесте", testSummary.getStudentsAbsent());

        // Рассчитываем процент присутствия
        double attendancePercentage = 0.0;
        if (testSummary.getStudentsPresent() != null && testSummary.getClassSize() != null &&
                testSummary.getClassSize() > 0) {
            attendancePercentage = (testSummary.getStudentsPresent() * 100.0) / testSummary.getClassSize();
        }
        startRow = addStatisticRow(sheet, startRow, "Процент присутствия",
                String.format("%.1f%%", attendancePercentage));

        startRow++; // Пустая строка

        startRow = addStatisticRow(sheet, startRow, "Количество заданий", testSummary.getTaskCount());
        startRow = addStatisticRow(sheet, startRow, "Максимальный балл за тест", testSummary.getMaxTotalScore());

        // Средний балл
        Double averageScore = testSummary.getAverageScore();
        if (averageScore == null && testSummary.getAverageScore() != null) {
            averageScore = testSummary.getAverageScore();
        }
        startRow = addStatisticRow(sheet, startRow, "Средний балл по классу",
                String.format("%.2f", averageScore != null ? averageScore : 0.0));

        // Процент выполнения
        double successPercentage = 0.0;
        if (testSummary.getMaxTotalScore() != null && testSummary.getMaxTotalScore() > 0 &&
                averageScore != null) {
            successPercentage = (averageScore * 100.0) / testSummary.getMaxTotalScore();
        }
        startRow = addStatisticRow(sheet, startRow, "Процент выполнения теста",
                String.format("%.1f%%", successPercentage));

        return startRow + 1; // Пустая строка в конце
    }

    /**
     * Создает результаты студентов
     */
    private int createStudentsResults(Workbook workbook, Sheet sheet,
                                      List<StudentDetailedResultDto> studentResults,
                                      Map<Integer, TaskStatisticsDto> taskStatistics, int startRow) {

        if (studentResults == null || studentResults.isEmpty()) {
            return startRow;
        }

        // Заголовок раздела
        Row headerRow = sheet.createRow(startRow++);
        headerRow.createCell(0).setCellValue("РЕЗУЛЬТАТЫ СТУДЕНТОВ");
        headerRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Определяем максимальное количество заданий
        int maxTaskNumber = taskStatistics != null && !taskStatistics.isEmpty() ?
                Collections.max(taskStatistics.keySet()) : 0;

        // Если статистики нет, определяем из данных студентов
        if (maxTaskNumber == 0) {
            for (StudentDetailedResultDto student : studentResults) {
                if (student.getTaskScores() != null && !student.getTaskScores().isEmpty()) {
                    int studentMax = Collections.max(student.getTaskScores().keySet());
                    maxTaskNumber = Math.max(maxTaskNumber, studentMax);
                }
            }
        }

        // Минимум 10 заданий для корректного отображения
        maxTaskNumber = Math.max(maxTaskNumber, 10);

        // Создаем заголовки столбцов
        Row columnHeaderRow = sheet.createRow(startRow++);
        String[] headers = new String[6 + maxTaskNumber];
        headers[0] = "№";
        headers[1] = "ФИО";
        headers[2] = "Присутствие";
        headers[3] = "Вариант";
        headers[4] = "Общий балл";
        headers[5] = "% выполнения";

        for (int i = 1; i <= maxTaskNumber; i++) {
            headers[5 + i] = "№" + i;
        }

        for (int i = 0; i < headers.length; i++) {
            Cell cell = columnHeaderRow.createCell(i);
            cell.setCellValue(headers[i]);
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

            // Общий балл
            if (student.getTotalScore() != null) {
                row.createCell(4).setCellValue(student.getTotalScore());
            } else {
                row.createCell(4).setCellValue(0);
            }

            // Процент выполнения
            Cell percentCell = row.createCell(5);
            if (student.getPercentageScore() != null) {
                percentCell.setCellValue(student.getPercentageScore() / 100.0);
            } else {
                percentCell.setCellValue(0);
            }
            percentCell.setCellStyle(createPercentStyle(workbook));

            // Баллы за задания
            if (student.getTaskScores() != null) {
                for (int taskNum = 1; taskNum <= maxTaskNumber; taskNum++) {
                    Integer score = student.getTaskScores().get(taskNum);
                    Cell scoreCell = row.createCell(5 + taskNum);
                    if (score != null) {
                        scoreCell.setCellValue(score);
                    } else {
                        scoreCell.setCellValue(0);
                    }
                    scoreCell.setCellStyle(createCenteredStyle(workbook));
                }
            } else {
                // Заполняем нулями
                for (int taskNum = 1; taskNum <= maxTaskNumber; taskNum++) {
                    Cell scoreCell = row.createCell(5 + taskNum);
                    scoreCell.setCellValue(0);
                    scoreCell.setCellStyle(createCenteredStyle(workbook));
                }
            }
        }

        return startRow + 1; // Пустая строка в конце
    }

    /**
     * Создает анализ по заданиям
     */
    private int createTaskAnalysis(Workbook workbook, Sheet sheet,
                                   Map<Integer, TaskStatisticsDto> taskStatistics, int startRow) {

        if (taskStatistics == null || taskStatistics.isEmpty()) {
            return startRow;
        }

        // Заголовок раздела
        Row headerRow = sheet.createRow(startRow++);
        headerRow.createCell(0).setCellValue("АНАЛИЗ ПО ЗАДАНИЯМ");
        headerRow.getCell(0).setCellStyle(createSectionHeaderStyle(workbook));

        // Заголовки таблицы
        Row columnHeaderRow = sheet.createRow(startRow++);
        String[] headers = {
                "№ задания", "Макс. балл", "Полностью", "Частично", "Не справилось",
                "% выполнения", "Распределение баллов"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = columnHeaderRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        // Сортируем задания по номеру
        List<Integer> sortedTaskNumbers = new ArrayList<>(taskStatistics.keySet());
        Collections.sort(sortedTaskNumbers);

        // Заполняем данные
        for (Integer taskNumber : sortedTaskNumbers) {
            TaskStatisticsDto stats = taskStatistics.get(taskNumber);
            if (stats == null) continue;

            Row row = sheet.createRow(startRow++);

            row.createCell(0).setCellValue("№" + taskNumber);
            row.createCell(1).setCellValue(stats.getMaxScore());
            row.createCell(2).setCellValue(stats.getFullyCompletedCount());
            row.createCell(3).setCellValue(stats.getPartiallyCompletedCount());
            row.createCell(4).setCellValue(stats.getNotCompletedCount());

            Cell percentCell = row.createCell(5);
            percentCell.setCellValue(stats.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(createPercentStyle(workbook));

            // Распределение баллов
            if (stats.getScoreDistribution() != null && !stats.getScoreDistribution().isEmpty()) {
                String distribution = formatScoreDistribution(stats.getScoreDistribution());
                row.createCell(6).setCellValue(distribution);
            } else {
                row.createCell(6).setCellValue("-");
            }
        }

        return startRow + 1; // Пустая строка в конце
    }

    // ============ ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ============

    private int addStatisticRow(Sheet sheet, int rowNum, String label, Object value) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(label);
        if (value != null) {
            row.createCell(1).setCellValue(value.toString());
        } else {
            row.createCell(1).setCellValue("");
        }
        return rowNum + 1;
    }

    private void setCellValue(Row row, int column, Object value) {
        if (value != null) {
            if (value instanceof Number) {
                row.createCell(column).setCellValue(((Number) value).doubleValue());
            } else {
                row.createCell(column).setCellValue(value.toString());
            }
        }
    }

    private String formatScoreDistribution(Map<Integer, Integer> distribution) {
        if (distribution == null || distribution.isEmpty()) {
            return "";
        }

        List<String> parts = new ArrayList<>();
        for (Map.Entry<Integer, Integer> entry : distribution.entrySet()) {
            parts.add(entry.getKey() + " баллов: " + entry.getValue());
        }
        return String.join("; ", parts);
    }

    // ============ СТИЛИ ============

    private CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createSubtitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setItalic(true);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
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
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
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
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createPercentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createCenteredStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    // ============ СОХРАНЕНИЕ ФАЙЛА ============

    private File saveWorkbookToFile(Workbook workbook, TestSummaryDto testSummary) throws IOException {
        // Создаем директорию для отчетов, если её нет
        if (!Files.exists(REPORTS_DIR)) {
            Files.createDirectories(REPORTS_DIR);
        }

        // Формируем имя файла
        String fileName = String.format("Детальный_отчет_%s_%s_%s.xlsx",
                testSummary.getSubject(),
                testSummary.getClassName(),
                testSummary.getTestDate().format(DateTimeFormatter.ofPattern("yyyyMMdd")));

        File file = REPORTS_DIR.resolve(fileName).toFile();

        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }

        log.info("✅ Детальный отчет сохранен: {}", file.getAbsolutePath());
        return file;
    }

    /**
     * Удаленные методы, которые вызывались с неправильными параметрами
     * (их можно удалить, так как они дублируют функциональность)
     */

    /**
     * Создает уникальное имя листа для рабочей книги
     */
    public String createUniqueSheetName(XSSFWorkbook workbook, TestSummaryDto testSummary) {
        String baseName = String.format("%s_%s_%s",
                testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9_]", "").substring(0,
                        Math.min(12, testSummary.getSubject().length())),
                testSummary.getClassName().replaceAll("[^a-zA-Zа-яА-Я0-9_]", ""),
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

    /**
     * Создает безопасное имя листа
     */
    public String createSafeSheetName(TestSummaryDto testSummary) {
        String sheetName = String.format("%s_%s",
                testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9_]", ""),
                testSummary.getClassName().replaceAll("[^a-zA-Zа-яА-Я0-9_]", ""));

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        return sheetName;
    }

    /**
     * Создает детальный отчет на листе в существующей рабочей книге
     */
    public void createDetailReportOnSheet(XSSFWorkbook workbook,
                                          TestSummaryDto testSummary,
                                          List<StudentDetailedResultDto> studentResults,
                                          Map<Integer, TaskStatisticsDto> taskStatistics,
                                          String sheetName) {

        XSSFSheet sheet = workbook.createSheet(sheetName);
        int currentRow = 0;

        try {
            // 1. Заголовок отчета
            currentRow = createReportHeader(workbook, sheet, testSummary, currentRow);

            // 2. Общая статистика теста
            currentRow = createGeneralStatistics(workbook, sheet, testSummary, currentRow);

            // 3. Результаты студентов
            currentRow = createStudentsResults(workbook, sheet, studentResults, taskStatistics, currentRow);

            // 4. Анализ по заданиям (основная таблица)
            currentRow = createTaskAnalysis(workbook, sheet, taskStatistics, currentRow);

            // 5. Графики (если есть данные)
            if (taskStatistics != null && !taskStatistics.isEmpty() &&
                    studentResults != null && !studentResults.isEmpty()) {

                excelChartService.createCharts(workbook, sheet, testSummary, taskStatistics, currentRow);
            }

            // 6. Оптимизируем ширину колонок
            optimizeDetailReportColumns(sheet, taskStatistics != null ? taskStatistics.size() : 0);

            log.info("✅ Детальный отчет создан на листе '{}': {} строк", sheetName, currentRow);

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета", e);
            Row errorRow = sheet.createRow(currentRow);
            errorRow.createCell(0).setCellValue("Ошибка при создании отчета: " + e.getMessage());
        }
    }

    /**
     * Оптимизирует ширину колонок для детального отчета
     */
    private void optimizeDetailReportColumns(XSSFSheet sheet, int taskCount) {
        int columnsToAutoSize = 6 + taskCount; // Базовые колонки + задания
        columnsToAutoSize = Math.min(columnsToAutoSize, 50); // Ограничиваем

        for (int i = 0; i < columnsToAutoSize; i++) {
            sheet.autoSizeColumn(i);
            // Минимальная ширина
            if (sheet.getColumnWidth(i) < 1500) {
                sheet.setColumnWidth(i, 1500);
            }
            // Максимальная ширина
            if (sheet.getColumnWidth(i) > 5000) {
                sheet.setColumnWidth(i, 5000);
            }
        }
    }
}