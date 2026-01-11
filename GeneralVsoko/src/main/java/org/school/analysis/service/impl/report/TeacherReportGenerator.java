package org.school.analysis.service.impl.report;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Component
@Slf4j
public class TeacherReportGenerator extends ExcelReportBase {

    private final DetailReportGenerator detailReportGenerator;

    public TeacherReportGenerator(DetailReportGenerator detailReportGenerator) {
        this.detailReportGenerator = detailReportGenerator;
    }

    public File generateTeacherReportWithDetails(
            String teacherName,
            List<TestSummaryDto> teacherTests,
            List<TeacherTestDetailDto> teacherTestDetails,
            String school) {

        log.info("Генерация детального отчета для учителя: {} ({} тестов)",
                teacherName, teacherTests.size());

        try {
            Path reportsPath = createReportsFolder(school);

            String fileName = String.format("Отчет_учителя_%s.xlsx",
                    teacherName.replace(" ", "_"));

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                createTeacherSummarySheet(workbook, teacherName, teacherTests);

                for (TeacherTestDetailDto testDetail : teacherTestDetails) {
                    createTeacherTestDetailSheet(workbook, testDetail);
                }

                return saveWorkbook(workbook, reportsPath, fileName);
            }

        } catch (Exception e) {
            log.error("Ошибка при создании детального отчета учителя", e);
            return null;
        }
    }

    private void createTeacherSummarySheet(XSSFWorkbook workbook, String teacherName,
                                           List<TestSummaryDto> teacherTests) {

        Sheet sheet = workbook.createSheet("Сводка по тестам");

        // Заголовок
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue("Отчет по тестам учителя: " + teacherName);
        titleRow.getCell(0).setCellStyle(createTitleStyle(workbook));

        // Дата генерации
        Row dateRow = sheet.createRow(1);
        dateRow.createCell(0).setCellValue(
                "Отчет сгенерирован: " + LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE)
        );

        // Заголовки таблицы
        String[] headers = {
                "Предмет", "Класс", "Дата", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения"
        };

        Row headerRow = sheet.createRow(3);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }

        // Данные тестов
        int rowNum = 4;
        for (TestSummaryDto test : teacherTests) {
            Row row = sheet.createRow(rowNum++);

            // Базовые данные
            row.createCell(0).setCellValue(test.getSubject());
            row.createCell(1).setCellValue(test.getClassName());
            row.createCell(2).setCellValue(test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            row.createCell(3).setCellValue(test.getTestType());

            // Посещаемость
            setNumericCellValue(row, 4, test.getStudentsPresent());
            setNumericCellValue(row, 5, test.getStudentsAbsent());
            setNumericCellValue(row, 6, test.getClassSize());

            // Процент присутствия
            Cell attendanceCell = row.createCell(7);
            double attendance = test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0;
            attendanceCell.setCellValue(attendance);
            attendanceCell.setCellStyle(createPercentStyle(workbook));

            // Результаты
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

        // Итоговая строка
        if (!teacherTests.isEmpty()) {
            addTeacherSummaryRow(sheet, rowNum, teacherTests, workbook);
        }
    }

    private void addTeacherSummaryRow(Sheet sheet, int rowNum,
                                      List<TestSummaryDto> tests, Workbook workbook) {

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

        // Заполняем ячейки
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

    private void createTeacherTestDetailSheet(XSSFWorkbook workbook,
                                              TeacherTestDetailDto testDetail) {

        TestSummaryDto testSummary = testDetail.getTestSummary();
        if (testSummary == null) {
            log.warn("Пропускаем тест без основных данных");
            return;
        }

        // Создаем уникальное имя листа
        String sheetName = detailReportGenerator.createUniqueSheetName(workbook, testSummary);

        try {
            // Создаем детальный отчет на отдельном листе
            detailReportGenerator.createDetailReportOnSheet(
                    workbook, testSummary,
                    testDetail.getStudentResults(),
                    testDetail.getTaskStatistics(),
                    sheetName);

            log.info("✅ Лист для теста '{}' создан", sheetName);

        } catch (Exception e) {
            log.error("❌ Ошибка создания листа для теста {}: {}",
                    testSummary.getFileName(), e.getMessage(), e);
        }
    }
}