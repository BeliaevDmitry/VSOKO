package org.school.analysis.service.impl.report;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

@Component
@Slf4j
public class SummaryReportGenerator extends ExcelReportBase {

    public File generateSummaryReport(List<TestSummaryDto> tests) {
        log.info("Генерация сводного отчета для {} тестов", tests.size());

        try {
            Path reportsPath = createReportsFolder();

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Сводка по тестам");

                createHeader(sheet, workbook);
                fillData(sheet, workbook, tests);
                addStatisticsRow(sheet, workbook, tests);
                optimizeColumnWidths(sheet, 14);

                return saveWorkbook(workbook, reportsPath, "Свод всех работ.xlsx");
            }

        } catch (IOException e) {
            log.error("Ошибка при создании сводного отчета", e);
            throw new RuntimeException("Ошибка при создании отчета", e);
        }
    }

    private void createHeader(Sheet sheet, Workbook workbook) {
        // Заголовок
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("СВОДНЫЙ ОТЧЕТ ПО ВСЕМ ТЕСТАМ");
        titleCell.setCellStyle(createTitleStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 13));

        // Дата формирования
        Row dateRow = sheet.createRow(1);
        dateRow.createCell(0).setCellValue(
                "Дата формирования: " + LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE)
        );

        // Заголовки колонок
        String[] headers = {
                "Учебный год", "Предмет", "Класс", "Учитель", "Дата теста", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения", "Кол-во заданий", "Макс. балл"
        };

        Row headerRow = sheet.createRow(3);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(createTableHeaderStyle(workbook));
        }
    }

    private void fillData(Sheet sheet, Workbook workbook, List<TestSummaryDto> tests) {
        int rowNum = 4;

        for (TestSummaryDto test : tests) {
            Row row = sheet.createRow(rowNum++);

            // Базовые данные
            row.createCell(0).setCellValue(test.getACADEMIC_YEAR());
            row.createCell(1).setCellValue(test.getSubject());
            row.createCell(2).setCellValue(test.getClassName());
            row.createCell(3).setCellValue(test.getTeacher());
            row.createCell(4).setCellValue(
                    test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            row.createCell(5).setCellValue(test.getTestType());

            // Посещаемость
            setNumericCellValue(row, 6, test.getStudentsPresent());
            setNumericCellValue(row, 7, test.getStudentsAbsent());
            setNumericCellValue(row, 8, test.getClassSize());

            // Процент присутствия
            Cell attendanceCell = row.createCell(9);
            double attendance = test.getAttendancePercentage() != null ?
                    test.getAttendancePercentage() / 100.0 : 0.0;
            attendanceCell.setCellValue(attendance);
            attendanceCell.setCellStyle(createPercentStyle(workbook));

            // Результаты
            Cell avgScoreCell = row.createCell(10);
            avgScoreCell.setCellValue(test.getAverageScore() != null ?
                    test.getAverageScore() : 0.0);
            avgScoreCell.setCellStyle(createDecimalStyle(workbook));

            Cell successCell = row.createCell(11);
            double success = test.getSuccessPercentage() != null ?
                    test.getSuccessPercentage() / 100.0 : 0.0;
            successCell.setCellValue(success);
            successCell.setCellStyle(createPercentStyle(workbook));

            // Детали теста
            setNumericCellValue(row, 12, test.getTaskCount());
            setNumericCellValue(row, 13, test.getMaxTotalScore());
        }
    }

    private void addStatisticsRow(Sheet sheet, Workbook workbook, List<TestSummaryDto> tests) {
        if (tests.isEmpty()) return;

        int lastRow = sheet.getLastRowNum() + 2;
        Row summaryRow = sheet.createRow(lastRow);

        // Заголовок итогов
        summaryRow.createCell(0).setCellValue("ИТОГО:");
        summaryRow.getCell(0).setCellStyle(createSummaryStyle(workbook));
        summaryRow.createCell(1).setCellValue("Тестов: " + tests.size());

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
        Cell avgPresentCell = summaryRow.createCell(6);
        avgPresentCell.setCellValue(String.format("%.1f", avgPresent));

        Cell avgAbsentCell = summaryRow.createCell(7);
        avgAbsentCell.setCellValue(String.format("%.1f", avgAbsent));

        Cell avgClassSizeCell = summaryRow.createCell(8);
        avgClassSizeCell.setCellValue(String.format("%.1f", avgClassSize));

        Cell avgAttendanceCell = summaryRow.createCell(9);
        avgAttendanceCell.setCellValue(avgAttendance);
        avgAttendanceCell.setCellStyle(createPercentStyle(workbook));

        Cell avgScoreCell = summaryRow.createCell(10);
        avgScoreCell.setCellValue(String.format("%.2f", avgScore));

        Cell avgSuccessCell = summaryRow.createCell(11);
        avgSuccessCell.setCellValue(avgSuccess);
        avgSuccessCell.setCellStyle(createPercentStyle(workbook));
    }
}