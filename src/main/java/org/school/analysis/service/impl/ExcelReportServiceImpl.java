package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.ExcelReportService;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class ExcelReportServiceImpl implements ExcelReportService {

    private static final String FINAL_REPORT_FOLDER = AppConfig.FINAL_REPORT_FOLDER;

    @Override
    public File generateSummaryReport(List<TestSummaryDto> tests) {
        log.info("Генерация сводного отчета для {} тестов", tests.size());

        try {
            // Создаем папку для отчетов, если не существует
            Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
            if (!Files.exists(reportsPath)) {
                Files.createDirectories(reportsPath);
            }

            // Создаем workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            // Лист со сводной информацией
            XSSFSheet sheet = workbook.createSheet("Сводка по тестам");

            // ЗАГОЛОВКИ
            String[] headers = {
                    "Школа",
                    "Предмет",
                    "Класс",
                    "Дата теста",
                    "Тип теста",
                    "Учитель",
                    "Присутствовало",
                    "Отсутствовало",
                    "Всего в классе",
                    "% присутствия",
                    "Кол-во заданий теста",
                    "Макс. балл",
                    "Средний балл теста",
                    "Файл"
            };

            // Создаем строку с заголовками
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = createHeaderStyle(workbook);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
                sheet.setColumnWidth(i, 4000);
            }

            // Создаем стили
            CellStyle dataStyle = createDataStyle(workbook);
            CellStyle centeredStyle = createCenteredStyle(workbook);
            CellStyle dateStyle = createCenteredStyle(workbook);
            CellStyle numberStyle = createNumberStyle(workbook);
            CellStyle decimalStyle = createDecimalStyle(workbook);
            CellStyle percentStyle = createPercentStyle(workbook);

            int rowNum = 1;
            for (TestSummaryDto test : tests) {
                Row row = sheet.createRow(rowNum++);

                // 0. Школа
                Cell cell0 = row.createCell(0);
                cell0.setCellValue(test.getSchool() != null ? test.getSchool() : "ГБОУ №7");
                cell0.setCellStyle(dataStyle);

                // 1. Предмет
                Cell cell1 = row.createCell(1);
                cell1.setCellValue(test.getSubject() != null ? test.getSubject() : "");
                cell1.setCellStyle(dataStyle);

                // 2. Класс
                Cell cell2 = row.createCell(2);
                cell2.setCellValue(test.getClassName() != null ? test.getClassName() : "");
                cell2.setCellStyle(centeredStyle);

                // 3. Дата теста
                Cell cell3 = row.createCell(3);
                if (test.getTestDate() != null) {
                    cell3.setCellValue(test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
                } else {
                    cell3.setCellValue("");
                }
                cell3.setCellStyle(dateStyle);

                // 4. Тип теста
                Cell cell4 = row.createCell(4);
                cell4.setCellValue(test.getTestType() != null ? test.getTestType() : "");
                cell4.setCellStyle(dataStyle);

                // 5. Учитель
                Cell cell5 = row.createCell(5);
                cell5.setCellValue(test.getTeacher() != null ? test.getTeacher() : "");
                cell5.setCellStyle(dataStyle);

                // 6. Присутствовало
                Cell cell6 = row.createCell(6);
                cell6.setCellValue(test.getStudentsPresent() != null ? test.getStudentsPresent() : 0);
                cell6.setCellStyle(numberStyle);

                // 7. Отсутствовало
                Cell cell7 = row.createCell(7);
                cell7.setCellValue(test.getStudentsAbsent() != null ? test.getStudentsAbsent() : 0);
                cell7.setCellStyle(numberStyle);

                // 8. Всего в классе
                Cell cell8 = row.createCell(8);
                cell8.setCellValue(test.getClassSize() != null ? test.getClassSize() : 0);
                cell8.setCellStyle(numberStyle);

                // 9. % присутствия
                Cell cell9 = row.createCell(9);
                double attendancePercentage = test.getAttendancePercentage() != null ?
                        test.getAttendancePercentage() / 100.0 : 0.0;
                cell9.setCellValue(attendancePercentage);
                cell9.setCellStyle(percentStyle);

                // 10. Кол-во заданий теста
                Cell cell10 = row.createCell(10);
                cell10.setCellValue(test.getTaskCount() != null ? test.getTaskCount() : 0);
                cell10.setCellStyle(numberStyle);

                // 11. Макс. балл
                Cell cell11 = row.createCell(11);
                cell11.setCellValue(test.getMaxTotalScore() != null ? test.getMaxTotalScore() : 0);
                cell11.setCellStyle(numberStyle);

                // 12. Средний балл теста
                Cell cell12 = row.createCell(12);
                cell12.setCellValue(test.getAverageScore() != null ? test.getAverageScore() : 0.0);
                cell12.setCellStyle(decimalStyle);

                // 13. Файл
                Cell cell13 = row.createCell(13);
                cell13.setCellValue(test.getFileName() != null ? test.getFileName() : "");
                cell13.setCellStyle(dataStyle);
            }

            // Автонастройка ширины колонок
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Добавляем итоговую строку со средними значениями
            addSummaryRow(sheet, rowNum, tests);

            // Сохраняем файл
            String fileName = "Свод всех работ.xlsx";
            Path filePath = reportsPath.resolve(fileName);

            try (FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {
                workbook.write(outputStream);
            }

            workbook.close();

            log.info("Сводный отчет сохранен: {}", filePath);
            return filePath.toFile();

        } catch (IOException e) {
            log.error("Ошибка при создании сводного отчета", e);
            throw new RuntimeException("Ошибка при создании отчета", e);
        }
    }

    /**
     * Добавляет итоговую строку со средними значениями
     */
    private void addSummaryRow(XSSFSheet sheet, int startRow, List<TestSummaryDto> tests) {
        if (tests.isEmpty()) return;

        Row summaryRow = sheet.createRow(startRow);
        CellStyle summaryStyle = createSummaryStyle(sheet.getWorkbook());

        // Заголовок "Итого"
        Cell headerCell = summaryRow.createCell(0);
        headerCell.setCellValue("Итого:");
        headerCell.setCellStyle(summaryStyle);

        // Рассчитываем средние значения
        double avgStudentsPresent = tests.stream()
                .filter(t -> t.getStudentsPresent() != null)
                .mapToInt(TestSummaryDto::getStudentsPresent)
                .average()
                .orElse(0.0);

        double avgStudentsAbsent = tests.stream()
                .filter(t -> t.getStudentsAbsent() != null)
                .mapToInt(TestSummaryDto::getStudentsAbsent)
                .average()
                .orElse(0.0);

        double avgClassSize = tests.stream()
                .filter(t -> t.getClassSize() != null)
                .mapToInt(TestSummaryDto::getClassSize)
                .average()
                .orElse(0.0);

        double avgAttendancePercentage = tests.stream()
                .filter(t -> t.getAttendancePercentage() != null)
                .mapToDouble(TestSummaryDto::getAttendancePercentage)
                .average()
                .orElse(0.0) / 100.0; // Преобразуем в десятичное для Excel

        double avgTaskCount = tests.stream()
                .filter(t -> t.getTaskCount() != null)
                .mapToInt(TestSummaryDto::getTaskCount)
                .average()
                .orElse(0.0);

        double avgMaxScore = tests.stream()
                .filter(t -> t.getMaxTotalScore() != null)
                .mapToInt(TestSummaryDto::getMaxTotalScore)
                .average()
                .orElse(0.0);

        double avgAverageScore = tests.stream()
                .filter(t -> t.getAverageScore() != null)
                .mapToDouble(TestSummaryDto::getAverageScore)
                .average()
                .orElse(0.0);

        // Заполняем остальные ячейки
        for (int i = 1; i < 14; i++) {
            Cell cell = summaryRow.createCell(i);
            cell.setCellStyle(summaryStyle);

            switch (i) {
                case 1: // Предмет - пусто
                case 2: // Класс - пусто
                case 3: // Дата - пусто
                case 4: // Тип теста - пусто
                case 5: // Учитель - пусто
                case 13: // Файл - пусто
                    cell.setCellValue("");
                    break;
                case 6: // Присутствовало (среднее)
                    cell.setCellValue(Math.round(avgStudentsPresent * 100.0) / 100.0);
                    break;
                case 7: // Отсутствовало (среднее)
                    cell.setCellValue(Math.round(avgStudentsAbsent * 100.0) / 100.0);
                    break;
                case 8: // Всего в классе (среднее)
                    cell.setCellValue(Math.round(avgClassSize * 100.0) / 100.0);
                    break;
                case 9: // % присутствия (среднее)
                    cell.setCellValue(avgAttendancePercentage);
                    break;
                case 10: // Кол-во заданий (среднее)
                    cell.setCellValue(Math.round(avgTaskCount * 100.0) / 100.0);
                    break;
                case 11: // Макс. балл (среднее)
                    cell.setCellValue(Math.round(avgMaxScore * 100.0) / 100.0);
                    break;
                case 12: // Средний балл (среднее)
                    cell.setCellValue(Math.round(avgAverageScore * 100.0) / 100.0);
                    break;
                default:
                    cell.setCellValue("");
            }
        }
    }

    /**
     * Создает стиль для заголовков
     */
    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для обычных данных
     */
    private CellStyle createDataStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для выровненных по центру данных
     */
    private CellStyle createCenteredStyle(XSSFWorkbook workbook) {
        CellStyle style = createDataStyle(workbook);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * Создает стиль для числовых данных
     */
    private CellStyle createNumberStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0"));
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для десятичных чисел
     */
    private CellStyle createDecimalStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для процентов
     */
    private CellStyle createPercentStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Создает стиль для итоговой строки
     */
    private CellStyle createSummaryStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        return style;
    }

    @Override
    public File generateTestDetailReport(
            TestSummaryDto testSummary,
            List<StudentDetailedResultDto> studentResults,
            Map<Integer, TaskStatisticsDto> taskStatistics) {

        log.info("Генерация детального отчета для теста: {} - {}",
                testSummary.getSubject(), testSummary.getClassName());

        try {
            Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
            if (!Files.exists(reportsPath)) {
                Files.createDirectories(reportsPath);
            }

            String fileName = String.format("Детальный_отчет_%s_%s_%s.xlsx",
                    testSummary.getSubject(),
                    testSummary.getClassName(),
                    testSummary.getTestDate().format(DateTimeFormatter.ofPattern("ddMMyyyy")));
            Path filePath = reportsPath.resolve(fileName);

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                // 1. Лист с результатами студентов
                createStudentResultsSheet(workbook, testSummary, studentResults);

                // 2. Лист с анализом по заданиям
                createTaskAnalysisSheet(workbook, testSummary, taskStatistics);

                // 3. Лист со статистикой (опционально)
                createStatisticsSheet(workbook, testSummary, studentResults, taskStatistics);

                // Сохраняем файл
                try (FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {
                    workbook.write(outputStream);
                }
            }

            log.info("Детальный отчет сохранен: {}", filePath);
            return filePath.toFile();

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
            Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
            if (!Files.exists(reportsPath)) {
                Files.createDirectories(reportsPath);
            }

            String fileName = String.format("Отчет_учителя_%s_%s.xlsx",
                    teacherName.replace(" ", "_"),
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmm")));
            Path filePath = reportsPath.resolve(fileName);

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                // 1. Сводный лист по всем тестам
                createTeacherSummarySheet(workbook, teacherName, teacherTests);

                // 2. Лист по каждому тесту
                for (int i = 0; i < teacherTests.size(); i++) {
                    createTeacherTestSheet(workbook, teacherTests.get(i), teacherName, i);
                }

                // Сохраняем файл
                try (FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {
                    workbook.write(outputStream);
                }
            }

            log.info("Отчет учителя сохранен: {}", filePath);
            return filePath.toFile();

        } catch (Exception e) {
            log.error("Ошибка при создании отчета учителя", e);
            return null;
        }
    }

    // Методы для создания листов:

    private void createStudentResultsSheet(Workbook workbook, TestSummaryDto test,
                                           List<StudentDetailedResultDto> students) {
        Sheet sheet = workbook.createSheet("Результаты студентов");

        // Заголовки
        String[] headers = {"№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"};

        // Определяем максимальное количество заданий
        int maxTasks = students.stream()
                .map(StudentDetailedResultDto::getTaskScores)
                .filter(Objects::nonNull)
                .mapToInt(Map::size)
                .max()
                .orElse(0);

        // Добавляем заголовки для заданий
        String[] allHeaders = Arrays.copyOf(headers, headers.length + maxTasks);
        for (int i = 0; i < maxTasks; i++) {
            allHeaders[headers.length + i] = "Зад. " + (i + 1);
        }

        // Создаем таблицу
        createHeaderRow(sheet, allHeaders);

        // Заполняем данные
        int rowNum = 1;
        for (int i = 0; i < students.size(); i++) {
            StudentDetailedResultDto student = students.get(i);
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(i + 1); // №
            row.createCell(1).setCellValue(student.getFio());
            row.createCell(2).setCellValue(student.getPresence());
            row.createCell(3).setCellValue(student.getVariant() != null ? student.getVariant() : "");
            row.createCell(4).setCellValue(student.getTotalScore() != null ? student.getTotalScore() : 0);
            row.createCell(5).setCellValue(student.getPercentageScore() != null ?
                    student.getPercentageScore() / 100.0 : 0.0);

            // Баллы по заданиям
            Map<Integer, Integer> taskScores = student.getTaskScores();
            for (int taskNum = 1; taskNum <= maxTasks; taskNum++) {
                Integer score = taskScores != null ? taskScores.get(taskNum) : null;
                row.createCell(5 + taskNum).setCellValue(score != null ? score : 0);
            }
        }

        // Автонастройка ширины
        for (int i = 0; i < allHeaders.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createTaskAnalysisSheet(Workbook workbook, TestSummaryDto test,
                                         Map<Integer, TaskStatisticsDto> statistics) {
        Sheet sheet = workbook.createSheet("Анализ по заданиям");

        // Заголовки
        String[] headers = {
                "Задание",
                "Макс. балл",
                "Полностью",
                "Частично",
                "Не справилось",
                "% выполнения",
                "Распределение баллов"
        };

        createHeaderRow(sheet, headers);

        // Заполняем данные
        int rowNum = 1;
        for (TaskStatisticsDto stats : statistics.values()) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(stats.getTaskNumber());
            row.createCell(1).setCellValue(stats.getMaxScore());
            row.createCell(2).setCellValue(stats.getFullyCompletedCount());
            row.createCell(3).setCellValue(stats.getPartiallyCompletedCount());
            row.createCell(4).setCellValue(stats.getNotCompletedCount());

            // Процент выполнения (double, не может быть null)
            double completionPercentage = stats.getCompletionPercentage();
            row.createCell(5).setCellValue(completionPercentage / 100.0);

            // Строка с распределением баллов
            if (stats.getScoreDistribution() != null) {
                String distribution = stats.getScoreDistribution().entrySet().stream()
                        .sorted(Map.Entry.comparingByKey())
                        .map(e -> String.format("%d баллов: %d", e.getKey(), e.getValue()))
                        .collect(Collectors.joining("; "));
                row.createCell(6).setCellValue(distribution);
            }
        }

        // Автонастройка ширины
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createStatisticsSheet(Workbook workbook, TestSummaryDto test,
                                       List<StudentDetailedResultDto> students,
                                       Map<Integer, TaskStatisticsDto> taskStatistics) {
        Sheet sheet = workbook.createSheet("Общая статистика");

        // Заголовок
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue("Общая статистика теста");

        // Основные показатели
        int rowNum = 2;
        String[][] statsData = {
                {"Предмет", test.getSubject()},
                {"Класс", test.getClassName()},
                {"Дата", test.getTestDate().format(DateTimeFormatter.ofPattern("dd.MM.yyyy"))},
                {"Учитель", test.getTeacher()},
                {"Всего студентов", String.valueOf(test.getClassSize())},
                {"Присутствовало", String.valueOf(test.getStudentsPresent())},
                {"Отсутствовало", String.valueOf(test.getStudentsAbsent())},
                {"% присутствия", String.format("%.1f%%", test.getAttendancePercentage())},
                {"Средний балл", String.format("%.2f", test.getAverageScore())},
                {"% выполнения", String.format("%.1f%%", test.getSuccessPercentage())}
        };

        for (String[] rowData : statsData) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(rowData[0]);
            row.createCell(1).setCellValue(rowData[1]);
        }

        // Автонастройка ширины
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
    }

    private void createTeacherSummarySheet(Workbook workbook, String teacherName,
                                           List<TestSummaryDto> teacherTests) {
        Sheet sheet = workbook.createSheet("Сводка по тестам");

        // Заголовок
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue("Отчет по тестам учителя: " + teacherName);

        // Заголовки таблицы
        String[] headers = {
                "Предмет", "Класс", "Дата", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения"
        };

        createHeaderRow(sheet, headers, 2);

        // Данные
        int rowNum = 3;
        for (TestSummaryDto test : teacherTests) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(test.getSubject());
            row.createCell(1).setCellValue(test.getClassName());
            row.createCell(2).setCellValue(test.getTestDate().format(DateTimeFormatter.ofPattern("dd.MM.yyyy")));
            row.createCell(3).setCellValue(test.getTestType());
            row.createCell(4).setCellValue(test.getStudentsPresent());
            row.createCell(5).setCellValue(test.getStudentsAbsent());
            row.createCell(6).setCellValue(test.getClassSize());
            row.createCell(7).setCellValue(test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0);
            row.createCell(8).setCellValue(test.getAverageScore());
            row.createCell(9).setCellValue(test.getSuccessPercentage() != null ?
                    test.getSuccessPercentage() / 100.0 : 0.0);
        }

        // Автонастройка ширины
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createTeacherTestSheet(Workbook workbook, TestSummaryDto test,
                                        String teacherName, int sheetIndex) {
        // Ограничиваем имя листа (макс 31 символ)
        String sheetName = String.format("%s_%s",
                test.getSubject().substring(0, Math.min(10, test.getSubject().length())),
                test.getClassName());

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        // Если имя уже существует, добавляем индекс
        if (workbook.getSheet(sheetName) != null) {
            sheetName = String.format("%s_%d", sheetName, sheetIndex);
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // Заголовок
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue(
                String.format("%s - %s (%s)", test.getSubject(), test.getClassName(),
                        test.getTestDate().format(DateTimeFormatter.ofPattern("dd.MM.yyyy")))
        );
    }

    private void createHeaderRow(Sheet sheet, String[] headers) {
        createHeaderRow(sheet, headers, 0);
    }

    private void createHeaderRow(Sheet sheet, String[] headers, int startRow) {
        Row headerRow = sheet.createRow(startRow);
        CellStyle headerStyle = createHeaderStyle((XSSFWorkbook) sheet.getWorkbook());

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }
}