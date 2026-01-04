package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.dto.StudentTestResultDto;
import org.school.analysis.model.dto.TestResultsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.AnalysisService;
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
import java.util.*;

import static org.school.analysis.util.DateTimeFormatters.*;


@Service
@RequiredArgsConstructor
@Slf4j
public class ExcelReportServiceImpl implements ExcelReportService {

    private final AnalysisService analysisService;

    // Используем константу из AppConfig
    private static final String FINAL_REPORT_FOLDER = AppConfig.FINAL_REPORT_FOLDER;

    // Стили для Excel
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

    private CellStyle createCenteredStyle(XSSFWorkbook workbook) {
        CellStyle style = createDataStyle(workbook);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createPercentageStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

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

            // ЗАГОЛОВКИ ПО ТРЕБОВАНИЯМ:
            // Школа, Предмет, Класс, Дата теста, Тип теста, Учитель,
            // Кол-во учеников писавших, Кол-во учеников в классе, Кол-во заданий теста,
            // Макс. балл, Средний балл теста
            // В generateSummaryReport добавляем колонки:
            String[] headers = {
                    "Школа",
                    "Предмет",
                    "Класс",
                    "Дата теста",
                    "Тип теста",
                    "Учитель",
                    "Присутствовало",        // Новое
                    "Отсутствовало",         // Новое
                    "Всего в классе",        // Новое
                    "% присутствия",         // Новое (вычисляемое)
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
            CellStyle numberStyle = createCenteredStyle(workbook);
            CellStyle decimalStyle = createDecimalStyle(workbook); // Для среднего балла
            CellStyle percentStyle = createPercentStyle(workbook);

            int rowNum = 1;
            for (TestSummaryDto test : tests) {
                Row row = sheet.createRow(rowNum++);

                // 0. Школа
                Cell cell0 = row.createCell(0);
                cell0.setCellValue(test.getSchool() != null ? test.getSchool() : "");
                cell0.setCellStyle(dataStyle);

                // 1. Предмет
                Cell cell1 = row.createCell(1);
                cell1.setCellValue(test.getSubject());
                cell1.setCellStyle(dataStyle);

                // 2. Класс
                Cell cell2 = row.createCell(2);
                cell2.setCellValue(test.getClassName());
                cell2.setCellStyle(centeredStyle);

                // 3. Дата теста
                Cell cell3 = row.createCell(3);
                cell3.setCellValue(test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
                cell3.setCellStyle(dateStyle);

                // 4. Тип теста
                Cell cell4 = row.createCell(4);
                cell4.setCellValue(test.getTestType() != null ? test.getTestType() : "");
                cell4.setCellStyle(dataStyle);

                // 5. Учитель
                Cell cell5 = row.createCell(5);
                cell5.setCellValue(test.getTeacher() != null ? test.getTeacher() : "");
                cell5.setCellStyle(dataStyle);

                // 6. Присутствовало (из теста кто был)
                Cell cell6 = row.createCell(6);
                cell6.setCellValue(test.getStudentsPresent() != null ? test.getStudentsPresent() : 0);
                cell6.setCellStyle(numberStyle);

                // 7. Отсутствовало (из теста кто не был)
                Cell cell7 = row.createCell(7);
                cell7.setCellValue(test.getStudentsAbsent() != null ? test.getStudentsAbsent() : 0);
                cell7.setCellStyle(numberStyle);

                // 8. Всего учеников в классе (из справочника)
                Cell cell8 = row.createCell(8);
                cell8.setCellValue(test.getClassSize() != null ? test.getClassSize() : 0);
                cell8.setCellStyle(numberStyle);

                // 9. % присутствия (вычисляемое)
                Cell cell9 = row.createCell(9);
                double attendanceRate = 0.0; // Переименуем для ясности
                if (test.getStudentsPresent() != null && test.getClassSize() != null
                        && test.getClassSize() > 0) {
                    attendanceRate = test.getStudentsPresent() * 1.0 / test.getClassSize();
                    // attendanceRate будет между 0.0 и 1.0
                }
                cell9.setCellValue(attendanceRate); // 0.7777, а не 77.777
                cell9.setCellStyle(percentStyle); // Новый стиль для процентов

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
                if (test.getAverageScore() != null) {
                    cell12.setCellValue(test.getAverageScore());
                } else {
                    cell12.setCellValue(0);
                }
                cell12.setCellStyle(decimalStyle);

                // 13. Файл (дополнительно)
                Cell cell13 = row.createCell(13);
                cell13.setCellValue(test.getFileName() != null ? test.getFileName() : "");
                cell13.setCellStyle(dataStyle);
            }
            // Автонастройка ширины колонок
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Добавляем итоговую строку со средними значениями
            addSummaryRow(sheet, rowNum, tests, headers.length);

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
     * Создает стиль для десятичных чисел (средний балл)
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
     * Добавляет итоговую строку со средними значениями
     */
    private void addSummaryRow(XSSFSheet sheet, int startRow, List<TestSummaryDto> tests, int columnCount) {
        if (tests.isEmpty()) return;

        Row summaryRow = sheet.createRow(startRow);

        // Заголовок "Итого"
        Cell headerCell = summaryRow.createCell(0);
        headerCell.setCellValue("Итого:");

        // Рассчитываем средние значения
        double avgStudentsCount = tests.stream()
                .filter(t -> t.getStudentsCount() != null)
                .mapToInt(TestSummaryDto::getStudentsCount)
                .average()
                .orElse(0.0);

        double avgClassSize = tests.stream()
                .filter(t -> t.getClassSize() != null)
                .mapToInt(TestSummaryDto::getClassSize)
                .average()
                .orElse(0.0);

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

        // Заполняем ячейки
        // Школа - пусто
        summaryRow.createCell(1).setCellValue("");
        // Предмет - пусто
        summaryRow.createCell(2).setCellValue("");
        // Класс - пусто
        summaryRow.createCell(3).setCellValue("");
        // Дата теста - пусто
        summaryRow.createCell(4).setCellValue("");
        // Тип теста - пусто
        summaryRow.createCell(5).setCellValue("");
        // Учитель - пусто
        // Кол-во учеников писавших (среднее)
        Cell cell6 = summaryRow.createCell(6);
        cell6.setCellValue(Math.round(avgStudentsCount * 100.0) / 100.0);
        // Кол-во учеников в классе (среднее)
        Cell cell7 = summaryRow.createCell(7);
        cell7.setCellValue(Math.round(avgClassSize * 100.0) / 100.0);
        // Кол-во заданий теста (среднее)
        Cell cell8 = summaryRow.createCell(8);
        cell8.setCellValue(Math.round(avgTaskCount * 100.0) / 100.0);
        // Макс. балл (среднее)
        Cell cell9 = summaryRow.createCell(9);
        cell9.setCellValue(Math.round(avgMaxScore * 100.0) / 100.0);
        // Средний балл теста (среднее из средних)
        Cell cell10 = summaryRow.createCell(10);
        cell10.setCellValue(Math.round(avgAverageScore * 100.0) / 100.0);
        // Файл - пусто
        summaryRow.createCell(11).setCellValue("");

        // Стиль для итоговой строки
        CellStyle summaryStyle = createSummaryStyle(sheet.getWorkbook());
        for (int i = 0; i < columnCount; i++) {
            Cell cell = summaryRow.getCell(i);
            if (cell != null) {
                cell.setCellStyle(summaryStyle);
            }
        }
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
    public File generateTeacherReport(TestResultsDto testResults, boolean includeChart) {
        if (testResults == null || testResults.getTestSummary() == null) {
            log.error("Нет данных для генерации отчета");
            return null;
        }

        TestSummaryDto summary = testResults.getTestSummary();
        log.info("Генерация отчета для учителя: {} {} {}",
                summary.getSubject(),
                summary.getClassName(),
                summary.getTestDate());

        try {
            // Создаем папку для отчетов, если не существует
            Path reportsPath = Paths.get(FINAL_REPORT_FOLDER);
            if (!Files.exists(reportsPath)) {
                Files.createDirectories(reportsPath);
            }

            // Создаем workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            // 1. Лист с результатами учеников
            XSSFSheet resultsSheet = workbook.createSheet("Результаты учеников");

            // Заголовок отчета
            createReportHeader(workbook, resultsSheet, testResults);

            // Заголовки таблицы
            String[] headers = {
                    "№", "ФИО ученика", "Присутствие", "Вариант",
                    "Балл", "Процент", "Место в классе"
            };

            Row headerRow = resultsSheet.createRow(5);
            CellStyle headerStyle = createHeaderStyle(workbook);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // Заполняем данные
            CellStyle dataStyle = createDataStyle(workbook);
            CellStyle centeredStyle = createCenteredStyle(workbook);
            CellStyle percentageStyle = createPercentageStyle(workbook);

            int rowNum = 6;
            int studentNumber = 1;

            if (testResults.getStudentResults() != null) {
                for (StudentTestResultDto student : testResults.getStudentResults()) {
                    Row row = resultsSheet.createRow(rowNum++);

                    // №
                    Cell cell0 = row.createCell(0);
                    cell0.setCellValue(studentNumber++);
                    cell0.setCellStyle(centeredStyle);

                    // ФИО ученика
                    Cell cell1 = row.createCell(1);
                    cell1.setCellValue(student.getFio() != null ? student.getFio() : "");
                    cell1.setCellStyle(dataStyle);

                    // Присутствие
                    Cell cell2 = row.createCell(2);
                    cell2.setCellValue(student.getPresence() != null ? student.getPresence() : "");
                    cell2.setCellStyle(centeredStyle);

                    // Вариант
                    Cell cell3 = row.createCell(3);
                    cell3.setCellValue(student.getVariant() != null ? student.getVariant() : "");
                    cell3.setCellStyle(centeredStyle);

                    // Балл
                    Cell cell4 = row.createCell(4);
                    cell4.setCellValue(student.getTotalScore() != null ? student.getTotalScore() : 0);
                    cell4.setCellStyle(centeredStyle);

                    // Процент
                    Cell cell5 = row.createCell(5);
                    if (student.getPercentageScore() != null) {
                        cell5.setCellValue(student.getPercentageScore() / 100.0);
                    } else {
                        cell5.setCellValue(0);
                    }
                    cell5.setCellStyle(percentageStyle);

                    // Место в классе
                    Cell cell6 = row.createCell(6);
                    cell6.setCellValue(student.getPositionInClass() != null ? student.getPositionInClass() : 0);
                    cell6.setCellStyle(centeredStyle);
                }
            }

            // Добавляем итоговую статистику
            addSummaryStatistics(workbook, resultsSheet, rowNum, testResults);

            // Автонастройка ширины колонок
            for (int i = 0; i < headers.length; i++) {
                resultsSheet.autoSizeColumn(i);
            }

            // 2. Лист с диаграммой (если нужно)
            if (includeChart) {
                try {
                    XSSFSheet chartSheet = workbook.createSheet("Анализ заданий");
                    createTaskAnalysisChart(workbook, chartSheet, testResults);
                } catch (Exception e) {
                    log.warn("Не удалось создать диаграмму: {}", e.getMessage());
                }
            }

            // Сохраняем файл
            String fileName = String.format("%s_%s_%s_%s.xlsx",
                    summary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9]", "_"),
                    summary.getClassName().replaceAll("[^a-zA-Zа-яА-Я0-9]", "_"),
                    summary.getTestDate().format(FILE_SIMPLE),
                    LocalDateTime.now().format(FILE_WITH_TIME));

            Path filePath = reportsPath.resolve(fileName);

            try (FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {
                workbook.write(outputStream);
            }

            workbook.close();

            log.info("Отчет для учителя сохранен: {}", filePath);
            return filePath.toFile();

        } catch (IOException e) {
            log.error("Ошибка при создании отчета для учителя", e);
            throw new RuntimeException("Ошибка при создании отчета", e);
        }
    }

    @Override
    public File generateTaskAnalysisReport(TestResultsDto testResults) {
        log.info("Генерация отчета с анализом заданий");

        try {
            // Для этого отчета нужны детальные данные по заданиям
            // Пока создадим упрощенную версию
            return generateTeacherReport(testResults, true);

        } catch (Exception e) {
            log.error("Ошибка при создании отчета с анализом заданий", e);
            throw new RuntimeException("Ошибка при создании отчета", e);
        }
    }

    @Override
    public List<File> generateAllReports(List<TestSummaryDto> tests) {
        log.info("Генерация всех отчетов для {} тестов", tests.size());

        List<File> generatedFiles = new ArrayList<>();

        try {
            // 1. Генерируем сводный отчет
            File summaryReport = generateSummaryReport(tests);
            if (summaryReport != null) {
                generatedFiles.add(summaryReport);
            }

            // 2. Генерируем детальные отчеты для каждого теста
            for (TestSummaryDto test : tests) {
                try {
                    TestResultsDto testResults = analysisService.getTestResults(
                            test.getSubject(),
                            test.getClassName(),
                            test.getTestDate(),
                            test.getTeacher()
                    );

                    if (testResults != null) {
                        File teacherReport = generateTeacherReport(testResults, true);
                        if (teacherReport != null) {
                            generatedFiles.add(teacherReport);
                        }
                    }
                } catch (Exception e) {
                    log.error("Ошибка при генерации отчета для теста {} {} {}: {}",
                            test.getSubject(), test.getClassName(), test.getTestDate(),
                            e.getMessage());
                }
            }

            log.info("Сгенерировано {} файлов отчетов", generatedFiles.size());
            return generatedFiles;

        } catch (Exception e) {
            log.error("Ошибка при генерации всех отчетов", e);
            throw new RuntimeException("Ошибка при создании отчетов", e);
        }
    }

    /**
     * Создает заголовок отчета
     */
    private void createReportHeader(XSSFWorkbook workbook, XSSFSheet sheet, TestResultsDto testResults) {
        TestSummaryDto summary = testResults.getTestSummary();

        // Объединяем ячейки для заголовка
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 6));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));

        // Стиль для заголовка
        CellStyle titleStyle = workbook.createCellStyle();
        XSSFFont titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 16);
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);

        CellStyle subtitleStyle = workbook.createCellStyle();
        XSSFFont subtitleFont = workbook.createFont();
        subtitleFont.setFontHeightInPoints((short) 12);
        subtitleStyle.setFont(subtitleFont);
        subtitleStyle.setAlignment(HorizontalAlignment.CENTER);

        // Заголовок
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(String.format("Отчет по тесту: %s", summary.getSubject()));
        titleCell.setCellStyle(titleStyle);

        // Подзаголовки
        Row row1 = sheet.createRow(1);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue(String.format("Класс: %s", summary.getClassName()));
        cell1.setCellStyle(subtitleStyle);

        Row row2 = sheet.createRow(2);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(String.format("Дата: %s", summary.getTestDate().format(DISPLAY_DATE)));
        cell2.setCellStyle(subtitleStyle);

        Row row3 = sheet.createRow(3);
        Cell cell3 = row3.createCell(0);
        cell3.setCellValue(String.format("Учитель: %s", summary.getTeacher()));
        cell3.setCellStyle(subtitleStyle);
    }

    /**
     * Добавляет итоговую статистику
     */
    private void addSummaryStatistics(XSSFWorkbook workbook, XSSFSheet sheet, int startRow, TestResultsDto testResults) {
        TestSummaryDto summary = testResults.getTestSummary();
        List<StudentTestResultDto> students = testResults.getStudentResults();

        if (students == null || students.isEmpty()) {
            return;
        }

        // Пропускаем одну строку
        int rowNum = startRow + 1;

        // Стиль для статистики
        CellStyle statHeaderStyle = createHeaderStyle(workbook);
        CellStyle statDataStyle = createDataStyle(workbook);
        CellStyle statCenteredStyle = createCenteredStyle(workbook);

        // Заголовок статистики
        Row statHeaderRow = sheet.createRow(rowNum++);
        sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 6));
        Cell statHeaderCell = statHeaderRow.createCell(0);
        statHeaderCell.setCellValue("Статистика по тесту");
        statHeaderCell.setCellStyle(statHeaderStyle);

        // Рассчитываем статистику
        double avgScore = students.stream()
                .filter(s -> s.getTotalScore() != null)
                .mapToInt(StudentTestResultDto::getTotalScore)
                .average()
                .orElse(0.0);

        int maxScore = students.stream()
                .filter(s -> s.getTotalScore() != null)
                .mapToInt(StudentTestResultDto::getTotalScore)
                .max()
                .orElse(0);

        int minScore = students.stream()
                .filter(s -> s.getTotalScore() != null)
                .mapToInt(StudentTestResultDto::getTotalScore)
                .min()
                .orElse(0);

        double avgPercentage = students.stream()
                .filter(s -> s.getPercentageScore() != null)
                .mapToDouble(StudentTestResultDto::getPercentageScore)
                .average()
                .orElse(0.0);

        long goodStudents = students.stream()
                .filter(s -> s.getPercentageScore() != null && s.getPercentageScore() >= 80)
                .count();

        long badStudents = students.stream()
                .filter(s -> s.getPercentageScore() != null && s.getPercentageScore() < 50)
                .count();

        // Данные статистики
        Object[][] statData = {
                {"Всего учеников:", students.size()},
                {"Средний балл:", String.format("%.2f", avgScore)},
                {"Максимальный балл:", maxScore},
                {"Минимальный балл:", minScore},
                {"Процент выполнения:", String.format("%.1f%%", avgPercentage)},
                {"Выполнили на 80% и выше:", goodStudents},
                {"Выполнили менее 50%:", badStudents}
        };

        // Заполняем статистику
        for (Object[] statRow : statData) {
            Row row = sheet.createRow(rowNum++);

            Cell labelCell = row.createCell(0);
            labelCell.setCellValue(statRow[0].toString());
            labelCell.setCellStyle(statHeaderStyle);

            Cell valueCell = row.createCell(1);
            valueCell.setCellValue(statRow[1].toString());
            valueCell.setCellStyle(statCenteredStyle);

            // Объединяем оставшиеся ячейки для аккуратного вида
            sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 2, 6));
        }
    }

    /**
     * Создает диаграмму анализа заданий
     */
    private void createTaskAnalysisChart(XSSFWorkbook workbook, XSSFSheet sheet, TestResultsDto testResults) {
        // В реальной реализации здесь будет логика анализа заданий
        // Пока создадим заглушку

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Анализ заданий");

        Row infoRow = sheet.createRow(1);
        Cell infoCell = infoRow.createCell(0);
        infoCell.setCellValue("Для детального анализа заданий необходимы данные из таблицы student_task_scores");

        // TODO: Реализовать сбор статистики по заданиям и построение диаграммы
        // Для этого нужно получать данные из таблицы student_task_scores
    }

    private CellStyle createPercentStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0%")); // Формат с 1 знаком после запятой
        // или "0%" если без десятичных
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
}