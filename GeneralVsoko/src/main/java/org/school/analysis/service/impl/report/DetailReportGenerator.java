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
import java.time.format.DateTimeFormatter;
import java.util.*;

import static org.school.analysis.config.AppConfig.*;

@Service
@RequiredArgsConstructor
@Slf4j
public class DetailReportGenerator extends ExcelReportBase {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    private final ExcelChartService excelChartService;

    /**
     * Создает отчёты по каждому тесту
     */
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

            // 2. Общая статистика (обновленная версия)
            currentRow = createGeneralStatistics(workbook, sheet, testSummary, currentRow);

            // 3. Результаты студентов
            currentRow = createStudentsResults(workbook, sheet, studentResults, taskStatistics, currentRow);

            // 4. Анализ по заданиям (основная таблица)
            currentRow = createTaskAnalysis(workbook, sheet, taskStatistics, currentRow);

            // 5. ГРАФИЧЕСКИЙ АНАЛИЗ
            currentRow = createGraphicalAnalysis(workbook, sheet, taskStatistics, currentRow);

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
     * Создает раздел с графиками (использует данные из таблицы "АНАЛИЗ ПО ЗАДАНИЯМ")
     */
    private int createGraphicalAnalysis(XSSFWorkbook workbook, XSSFSheet sheet,
                                        Map<Integer, TaskStatisticsDto> taskStatistics,
                                        int startRow) {

        log.debug("Создание графического анализа, стартовая строка: {}", startRow);

        // Пустая строка перед разделом
        startRow += 2;

        // ========== 1. ЗАГОЛОВОК "ГРАФИЧЕСКИЙ АНАЛИЗ" ==========
        Row sectionHeader = sheet.createRow(startRow++);
        Cell headerCell = sectionHeader.createCell(0); // Колонка A

        headerCell.setCellValue("ГРАФИЧЕСКИЙ АНАЛИЗ");

        // Стиль для заголовка раздела (совпадает с другими заголовками)
        CellStyle sectionHeaderStyle = workbook.createCellStyle();
        Font sectionFont = workbook.createFont();
        sectionFont.setBold(true);
        sectionFont.setFontHeightInPoints((short) 12);
        sectionFont.setColor(IndexedColors.WHITE.getIndex());
        sectionHeaderStyle.setFont(sectionFont);
        sectionHeaderStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        sectionHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sectionHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        sectionHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Границы для заголовка
        sectionHeaderStyle.setBorderTop(BorderStyle.THIN);
        sectionHeaderStyle.setBorderBottom(BorderStyle.THIN);
        sectionHeaderStyle.setBorderLeft(BorderStyle.THIN);
        sectionHeaderStyle.setBorderRight(BorderStyle.THIN);

        headerCell.setCellStyle(sectionHeaderStyle);

        // ОБЪЕДИНЕНИЕ A-P (16 столбцов: A=0, P=15)
        int mergeEndColumn = 15; // Колонка P (индекс 15) - должно совпадать с блоком "АНАЛИЗ ПО ЗАДАНИЯМ"
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, mergeEndColumn));

        // ========== 2. ПОДЗАГОЛОВОК (описание) ==========
        Row descriptionRow = sheet.createRow(startRow++);
        Cell descCell = descriptionRow.createCell(0);
        descCell.setCellValue("*Данные для графиков взяты из раздела 'АНАЛИЗ ПО ЗАДАНИЯМ'");

        // Стиль для описания
        CellStyle descStyle = workbook.createCellStyle();
        Font descFont = workbook.createFont();
        descFont.setItalic(true);
        descFont.setFontHeightInPoints((short) 10);
        descFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        descStyle.setFont(descFont);
        descStyle.setAlignment(HorizontalAlignment.LEFT);
        descCell.setCellStyle(descStyle);

        // Объединяем для подзаголовка тоже (A-P)
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, mergeEndColumn));

        startRow++; // Пустая строка после описания

        // ========== 3. ПРОВЕРКА ДАННЫХ ДЛЯ ГРАФИКОВ ==========
        if (taskStatistics == null || taskStatistics.isEmpty()) {
            Row noDataRow = sheet.createRow(startRow++);
            noDataRow.createCell(0).setCellValue("Нет данных для графиков");
            // Объединяем ячейки для сообщения
            sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, mergeEndColumn));
            return startRow;
        }

        // ========== 4. ПОИСК ТАБЛИЦЫ "АНАЛИЗ ПО ЗАДАНИЯМ" ==========
        int analysisTableStartRow = -1;
        for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    if ("АНАЛИЗ ПО ЗАДАНИЯМ".equals(cellValue)) {
                        analysisTableStartRow = rowNum;
                        break;
                    }
                }
            }
        }

        // ========== 5. СОЗДАНИЕ ГРАФИКОВ ==========
        if (analysisTableStartRow != -1) {
            // Определяем диапазон данных
            int dataStartRow = analysisTableStartRow + 2; // Пропускаем заголовок и заголовки столбцов

            // Создаем графики через ExcelChartService
            excelChartService.createChartsFromAnalysisTable(workbook, sheet,
                    dataStartRow, taskStatistics.size(), startRow);

            // Возвращаем следующую свободную строку
            return sheet.getLastRowNum() + 2;
        } else {
            // Если таблица не найдена, создаем графики из предоставленных данных
            log.info("Таблица анализ данных не найдена");

            // Создаем заглушку для теста
            Row placeholderRow = sheet.createRow(startRow++);
            placeholderRow.createCell(0).setCellValue("Графики будут созданы здесь на основе данных");
            sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, mergeEndColumn));

            return sheet.getLastRowNum() + 2;
        }
    }

    /**
     * Оптимизирует ширину столбцов по содержимому
     */
    private void autoSizeColumns(XSSFSheet sheet, int taskCount) {
        // Определяем общее количество столбцов
        int totalColumns = 6 + taskCount; // 6 базовых + задания
        totalColumns = Math.min(totalColumns, 60); // Ограничиваем разумным максимумом

        // Просто автонастраиваем ширину для каждого столбца
        for (int i = 0; i < totalColumns; i++) {
            sheet.autoSizeColumn(i);
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
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));

        // Подзаголовок
        Row subtitleRow = sheet.createRow(startRow++);
        Cell subtitleCell = subtitleRow.createCell(0);
        subtitleCell.setCellValue(
                String.format("Дата проведения: %s | Тип работы: %s",
                        testSummary.getTestDate().format(DATE_FORMATTER),
                        testSummary.getTestType())
        );
        subtitleCell.setCellStyle(createSubtitleStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));

        // Дополнительная информация
        if (testSummary.getTeacher() != null && !testSummary.getTeacher().isEmpty()) {
            Row teacherRow = sheet.createRow(startRow++);
            teacherRow.createCell(0).setCellValue("Учитель: " + testSummary.getTeacher());
            sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));
        }

        if (testSummary.getFileName() != null && !testSummary.getFileName().isEmpty()) {
            Row fileNameRow = sheet.createRow(startRow++);
            fileNameRow.createCell(0).setCellValue("Файл отчета: " + testSummary.getFileName());
            sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));
        }

        return startRow + 1; // Пустая строка
    }

    /**
     * Создает общую статистику теста (ФИНАЛЬНАЯ ИСПРАВЛЕННАЯ ВЕРСИЯ)
     */
    private int createGeneralStatistics(Workbook workbook, Sheet sheet,
                                        TestSummaryDto testSummary, int startRow) {

        // ========== 1. ЗАГОЛОВОК РАЗДЕЛА ==========
        // Заголовок должен быть в колонке B (индекс 1), а не в A
        Row headerRow = sheet.createRow(startRow);
        Cell headerCell = headerRow.createCell(1); // КОЛОНКА B

        headerCell.setCellValue("ОБЩАЯ СТАТИСТИКА ТЕСТА");

        // Стиль для заголовка раздела
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);

        headerCell.setCellStyle(headerStyle);

        // Объединяем B-C для заголовка
        sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 1, 3));

        startRow++; // Переходим к следующей строке

        // ========== 2. СОЗДАЕМ СТИЛИ ==========
        // Стиль для ОБЪЕДИНЕННЫХ ЯЧЕЕК B-C - с границами ВО ВСЕХ СТОРОНАХ
        CellStyle mergedCellStyle = workbook.createCellStyle();
        mergedCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        mergedCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // ВАЖНО: границы со ВСЕХ сторон для объединенной ячейки
        mergedCellStyle.setBorderTop(BorderStyle.THIN);
        mergedCellStyle.setBorderBottom(BorderStyle.THIN);
        mergedCellStyle.setBorderLeft(BorderStyle.THIN);  // Левая граница
        mergedCellStyle.setBorderRight(BorderStyle.THIN); // Правая граница

        mergedCellStyle.setAlignment(HorizontalAlignment.RIGHT);
        mergedCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font labelFont = workbook.createFont();
        labelFont.setBold(true);
        mergedCellStyle.setFont(labelFont);

        // Стиль для колонки D
        CellStyle valueStyle = workbook.createCellStyle();
        valueStyle.setBorderTop(BorderStyle.THIN);
        valueStyle.setBorderBottom(BorderStyle.THIN);
        valueStyle.setBorderLeft(BorderStyle.THIN);
        valueStyle.setBorderRight(BorderStyle.THIN);
        valueStyle.setAlignment(HorizontalAlignment.CENTER);
        valueStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Стиль для процентов
        CellStyle percentStyle = workbook.createCellStyle();
        percentStyle.cloneStyleFrom(valueStyle);
        DataFormat percentFormat = workbook.createDataFormat();
        percentStyle.setDataFormat(percentFormat.getFormat("0.0%"));

        // Стиль для десятичных
        CellStyle decimalStyle = workbook.createCellStyle();
        decimalStyle.cloneStyleFrom(valueStyle);
        DataFormat decimalFormat = workbook.createDataFormat();
        decimalStyle.setDataFormat(decimalFormat.getFormat("0.00"));

        // ========== 3. ДАННЫЕ СТАТИСТИКИ ==========
        Object[][] statisticsData = {
                {"Всего учеников в классе",
                        testSummary.getClassSize() != null ? testSummary.getClassSize() : 0,
                        "integer"},
                {"Присутствовало на тесте",
                        testSummary.getStudentsPresent() != null ? testSummary.getStudentsPresent() : 0,
                        "integer"},
                {"Отсутствовало на тесте",
                        testSummary.getStudentsAbsent() != null ? testSummary.getStudentsAbsent() : 0,
                        "integer"},
                {"Процент присутствия",
                        calculateAttendancePercentage(testSummary),
                        "percent"},
                {"Количество заданий",
                        testSummary.getTaskCount() != null ? testSummary.getTaskCount() : 0,
                        "integer"},
                {"Максимальный балл за тест",
                        testSummary.getMaxTotalScore() != null ? testSummary.getMaxTotalScore() : 0,
                        "integer"},
                {"Средний балл по классу",
                        testSummary.getAverageScore() != null ? testSummary.getAverageScore() : 0.0,
                        "decimal"},
                {"Процент выполнения теста",
                        calculateSuccessPercentage(testSummary),
                        "percent"}
        };

        // ========== 4. СОЗДАЕМ ТАБЛИЦУ ==========
        for (Object[] rowData : statisticsData) {
            Row row = sheet.createRow(startRow);

            // 1. Создаем ячейку в колонке B (это будет левая часть объединенной B-C)
            Cell mergedCell = row.createCell(1); // Колонка B
            mergedCell.setCellValue((String) rowData[0]);
            mergedCell.setCellStyle(mergedCellStyle);

            // 2. Создаем ПУСТУЮ ячейку в колонке C
            // Это важно для правильного отображения границ
            Cell cellC = row.createCell(2); // Колонка C
            // Устанавливаем ТОТ ЖЕ САМЫЙ стиль, что и для ячейки B
            cellC.setCellStyle(mergedCellStyle);

            // 3. Объединяем B и C
            sheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 1, 2));

            // 4. Создаем ячейку в колонке D (значение)
            Cell valueCell = row.createCell(3); // Колонка D
            Object value = rowData[1];
            String styleType = (String) rowData[2];

            if (value instanceof Number) {
                valueCell.setCellValue(((Number) value).doubleValue());

                switch (styleType) {
                    case "percent":
                        valueCell.setCellStyle(percentStyle);
                        break;
                    case "decimal":
                        valueCell.setCellStyle(decimalStyle);
                        break;
                    default:
                        valueCell.setCellStyle(valueStyle);
                        break;
                }
            }

            startRow++;
        }

        // ========== 5. НАСТРАИВАЕМ ШИРИНУ ==========
        sheet.setColumnWidth(1, 30 * 256);  // Ширина колонки B
        sheet.setColumnWidth(2, 5 * 256);   // Колонка C (узкая, но видимая)
        sheet.setColumnWidth(3, 15 * 256);  // Ширина колонки D

        return startRow + 1; // Пустая строка после блока
    }

    /**
     * Вспомогательный метод для расчета процента выполнения теста
     */
    private double calculateSuccessPercentage(TestSummaryDto testSummary) {
        if (testSummary.getAverageScore() != null &&
                testSummary.getMaxTotalScore() != null &&
                testSummary.getMaxTotalScore() > 0) {
            return testSummary.getAverageScore() / testSummary.getMaxTotalScore();
        }
        return 0.0;
    }

    /**
     * Вспомогательный метод для расчета процента присутствия
     */
    private double calculateAttendancePercentage(TestSummaryDto testSummary) {
        if (testSummary.getStudentsPresent() != null &&
                testSummary.getClassSize() != null &&
                testSummary.getClassSize() > 0) {
            return testSummary.getStudentsPresent() / (double) testSummary.getClassSize();
        }
        return 0.0;
    }

    /**
     * Создает результаты студентов с правильным оформлением
     */
    private int createStudentsResults(Workbook workbook, Sheet sheet,
                                      List<StudentDetailedResultDto> studentResults,
                                      Map<Integer, TaskStatisticsDto> taskStatistics, int startRow) {

        if (studentResults == null || studentResults.isEmpty()) {
            return startRow;
        }

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

        // ========== 1. ЗАГОЛОВОК РАЗДЕЛА ==========
        Row sectionHeader = sheet.createRow(startRow++);
        Cell headerCell = sectionHeader.createCell(0);
        headerCell.setCellValue("РЕЗУЛЬТАТЫ СТУДЕНТОВ");

        // Стиль для заголовка раздела
        CellStyle sectionHeaderStyle = workbook.createCellStyle();
        Font sectionFont = workbook.createFont();
        sectionFont.setBold(true);
        sectionFont.setFontHeightInPoints((short) 12);
        sectionFont.setColor(IndexedColors.WHITE.getIndex());
        sectionHeaderStyle.setFont(sectionFont);
        sectionHeaderStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        sectionHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sectionHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        sectionHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Границы для заголовка
        sectionHeaderStyle.setBorderTop(BorderStyle.THIN);
        sectionHeaderStyle.setBorderBottom(BorderStyle.THIN);
        sectionHeaderStyle.setBorderLeft(BorderStyle.THIN);
        sectionHeaderStyle.setBorderRight(BorderStyle.THIN);

        headerCell.setCellStyle(sectionHeaderStyle);

        // ОБЪЕДИНЕНИЕ: заголовок должен объединяться по всем колонкам с заданиями
        // Всего столбцов: 6 (базовые) + maxTaskNumber (задания)
        int totalColumns = 6 + maxTaskNumber; // №, ФИО, Присутствие, Вариант, Общий балл, % выполнения + задания
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, totalColumns - 1));

        // ========== 2. СОЗДАЕМ СТИЛИ ==========
        // Стиль для заголовков столбцов (жирный, с границами, по центру)
        CellStyle columnHeaderStyle = workbook.createCellStyle();
        Font columnHeaderFont = workbook.createFont();
        columnHeaderFont.setBold(true);
        columnHeaderFont.setFontHeightInPoints((short) 11);
        columnHeaderStyle.setFont(columnHeaderFont);
        columnHeaderStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        columnHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        columnHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        columnHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        columnHeaderStyle.setBorderTop(BorderStyle.THIN);
        columnHeaderStyle.setBorderBottom(BorderStyle.THIN);
        columnHeaderStyle.setBorderLeft(BorderStyle.THIN);
        columnHeaderStyle.setBorderRight(BorderStyle.THIN);

        // Стиль для данных (с границами, по центру)
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setWrapText(true); // Перенос текста

        // Стиль для данных с процентом
        CellStyle percentDataStyle = workbook.createCellStyle();
        percentDataStyle.cloneStyleFrom(dataStyle);
        DataFormat percentFormat = workbook.createDataFormat();
        percentDataStyle.setDataFormat(percentFormat.getFormat("0.0%"));

        // ========== 3. СОЗДАЕМ ЗАГОЛОВКИ СТОЛБЦОВ ==========
        Row columnHeaderRow = sheet.createRow(startRow++);

        // Базовые заголовки
        String[] baseHeaders = {"№", "ФИО", "Присутствие", "Вариант", "Общий балл", "% выполнения"};

        for (int i = 0; i < baseHeaders.length; i++) {
            Cell cell = columnHeaderRow.createCell(i);
            cell.setCellValue(baseHeaders[i]);
            cell.setCellStyle(columnHeaderStyle);
        }

        // Заголовки заданий
        for (int taskNum = 1; taskNum <= maxTaskNumber; taskNum++) {
            Cell cell = columnHeaderRow.createCell(baseHeaders.length + taskNum - 1);
            cell.setCellValue(Integer.toString(taskNum));
            cell.setCellStyle(columnHeaderStyle);
        }

        // ========== 4. ЗАПОЛНЯЕМ ДАННЫЕ СТУДЕНТОВ ==========
        for (int studentIndex = 0; studentIndex < studentResults.size(); studentIndex++) {
            StudentDetailedResultDto student = studentResults.get(studentIndex);
            Row row = sheet.createRow(startRow++);

            // Базовые данные студента (все по центру и с границами)

            // 1. №
            Cell numberCell = row.createCell(0);
            numberCell.setCellValue(studentIndex + 1);
            numberCell.setCellStyle(dataStyle);

            // 2. ФИО
            Cell fioCell = row.createCell(1);
            fioCell.setCellValue(student.getFio() != null ? student.getFio() : "");
            fioCell.setCellStyle(dataStyle);

            // 3. Присутствие
            Cell presenceCell = row.createCell(2);
            presenceCell.setCellValue(student.getPresence() != null ? student.getPresence() : "");
            presenceCell.setCellStyle(dataStyle);

            // 4. Вариант
            Cell variantCell = row.createCell(3);
            variantCell.setCellValue(student.getVariant() != null ? student.getVariant() : "");
            variantCell.setCellStyle(dataStyle);

            // 5. Общий балл (целое число)
            Cell totalScoreCell = row.createCell(4);
            if (student.getTotalScore() != null) {
                totalScoreCell.setCellValue(student.getTotalScore());
            } else {
                totalScoreCell.setCellValue(0);
            }
            totalScoreCell.setCellStyle(dataStyle);

            // 6. % выполнения (процентный формат)
            Cell percentCell = row.createCell(5);
            if (student.getPercentageScore() != null) {
                percentCell.setCellValue(student.getPercentageScore() / 100.0);
            } else {
                percentCell.setCellValue(0);
            }
            percentCell.setCellStyle(percentDataStyle);

            // 7. Баллы за задания (все по центру и с границами)
            for (int taskNum = 1; taskNum <= maxTaskNumber; taskNum++) {
                Cell scoreCell = row.createCell(5 + taskNum); // 5 базовых столбцов + номер задания

                if (student.getTaskScores() != null && student.getTaskScores().containsKey(taskNum)) {
                    Integer score = student.getTaskScores().get(taskNum);
                    scoreCell.setCellValue(score != null ? score : 0);
                } else {
                    scoreCell.setCellValue(0);
                }

                scoreCell.setCellStyle(dataStyle);
            }
        }

        // ========== 5. НАСТРАИВАЕМ ШИРИНУ СТОЛБЦОВ ==========
        sheet.setColumnWidth(0, 5 * 256);    // №
        sheet.setColumnWidth(1, 25 * 256);   // ФИО
        sheet.setColumnWidth(2, 15 * 256);   // Присутствие
        sheet.setColumnWidth(3, 10 * 256);   // Вариант
        sheet.setColumnWidth(4, 12 * 256);   // Общий балл
        sheet.setColumnWidth(5, 12 * 256);   // % выполнения

        // Ширина для столбцов с заданиями
        for (int taskNum = 1; taskNum <= maxTaskNumber; taskNum++) {
            sheet.setColumnWidth(5 + taskNum, 8 * 256); // Узкие столбцы для заданий
        }

        return startRow + 1; // Пустая строка после блока
    }

    /**
     * Создает анализ по заданиям с правильным оформлением
     */
    private int createTaskAnalysis(Workbook workbook, Sheet sheet,
                                   Map<Integer, TaskStatisticsDto> taskStatistics, int startRow) {

        if (taskStatistics == null || taskStatistics.isEmpty()) {
            return startRow;
        }

        // ========== 1. ЗАГОЛОВОК РАЗДЕЛА ==========
        // Заголовок "АНАЛИЗ ПО ЗАДАНИЯМ" должен объединять ячейки A-P
        Row sectionHeader = sheet.createRow(startRow++);
        Cell headerCell = sectionHeader.createCell(0); // Колонка A

        headerCell.setCellValue("АНАЛИЗ ПО ЗАДАНИЯМ");

        // Стиль для заголовка раздела
        CellStyle sectionHeaderStyle = workbook.createCellStyle();
        Font sectionFont = workbook.createFont();
        sectionFont.setBold(true);
        sectionFont.setFontHeightInPoints((short) 12);
        sectionFont.setColor(IndexedColors.WHITE.getIndex());
        sectionHeaderStyle.setFont(sectionFont);
        sectionHeaderStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        sectionHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sectionHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        sectionHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Границы для заголовка
        sectionHeaderStyle.setBorderTop(BorderStyle.THIN);
        sectionHeaderStyle.setBorderBottom(BorderStyle.THIN);
        sectionHeaderStyle.setBorderLeft(BorderStyle.THIN);
        sectionHeaderStyle.setBorderRight(BorderStyle.THIN);

        headerCell.setCellStyle(sectionHeaderStyle);

        // ОБЪЕДИНЕНИЕ A-P (16 столбцов: A=0, P=15)
        int mergeEndColumn = 15; // Колонка P (индекс 15)
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, mergeEndColumn));

        // ========== 2. СОЗДАЕМ СТИЛИ ==========
        // Стиль для заголовков столбцов
        CellStyle columnHeaderStyle = workbook.createCellStyle();
        Font columnHeaderFont = workbook.createFont();
        columnHeaderFont.setBold(true);
        columnHeaderFont.setFontHeightInPoints((short) 11);
        columnHeaderStyle.setFont(columnHeaderFont);
        columnHeaderStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        columnHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        columnHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        columnHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        columnHeaderStyle.setBorderTop(BorderStyle.THIN);
        columnHeaderStyle.setBorderBottom(BorderStyle.THIN);
        columnHeaderStyle.setBorderLeft(BorderStyle.THIN);
        columnHeaderStyle.setBorderRight(BorderStyle.THIN);

        // Стиль для данных
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setWrapText(true);

        // Стиль для процентов
        CellStyle percentStyle = workbook.createCellStyle();
        percentStyle.cloneStyleFrom(dataStyle);
        DataFormat percentFormat = workbook.createDataFormat();
        percentStyle.setDataFormat(percentFormat.getFormat("0.0%"));

        // Стиль для объединенной ячейки "Распределение баллов"
        CellStyle distributionStyle = workbook.createCellStyle();
        distributionStyle.setAlignment(HorizontalAlignment.LEFT); // Текст слева внутри объединенной ячейки
        distributionStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        distributionStyle.setBorderTop(BorderStyle.THIN);
        distributionStyle.setBorderBottom(BorderStyle.THIN);
        distributionStyle.setBorderLeft(BorderStyle.THIN);
        distributionStyle.setBorderRight(BorderStyle.THIN);
        distributionStyle.setWrapText(true);

        // ========== 3. СОЗДАЕМ ЗАГОЛОВКИ СТОЛБЦОВ ==========
        Row columnHeaderRow = sheet.createRow(startRow++);

        // Заголовки столбцов A-F
        String[] headers = {"№", "Макс. балл", "Полностью", "Частично", "Не справилось", "% выполнения"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = columnHeaderRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(columnHeaderStyle);
        }

        /// Заголовок "Распределение баллов" в столбце G (индекс 6)
        Cell distributionHeader = columnHeaderRow.createCell(6);
        distributionHeader.setCellValue("Распределение баллов");
        distributionHeader.setCellStyle(columnHeaderStyle);

// Создаём остальные ячейки H-P с ТЕМ ЖЕ стилем
        for (int col = 7; col <= mergeEndColumn; col++) {
            Cell emptyCell = columnHeaderRow.createCell(col);
            emptyCell.setCellStyle(columnHeaderStyle);
        }

// Теперь объединяем
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 6, mergeEndColumn));
        // ========== 4. ЗАПОЛНЯЕМ ДАННЫЕ ПО ЗАДАНИЯМ ==========
        // Сортируем задания по номеру
        List<Integer> sortedTaskNumbers = new ArrayList<>(taskStatistics.keySet());
        Collections.sort(sortedTaskNumbers);

        for (Integer taskNumber : sortedTaskNumbers) {
            TaskStatisticsDto stats = taskStatistics.get(taskNumber);
            if (stats == null) continue;

            Row row = sheet.createRow(startRow++);

            // 1. № задания
            Cell taskNumCell = row.createCell(0);
            taskNumCell.setCellValue("№" + taskNumber);
            taskNumCell.setCellStyle(dataStyle);

            // 2. Макс. балл
            Cell maxScoreCell = row.createCell(1);
            maxScoreCell.setCellValue(stats.getMaxScore());
            maxScoreCell.setCellStyle(dataStyle);

            // 3. Полностью
            Cell fullyCell = row.createCell(2);
            fullyCell.setCellValue(stats.getFullyCompletedCount());
            fullyCell.setCellStyle(dataStyle);

            // 4. Частично
            Cell partiallyCell = row.createCell(3);
            partiallyCell.setCellValue(stats.getPartiallyCompletedCount());
            partiallyCell.setCellStyle(dataStyle);

            // 5. Не справилось
            Cell notCompletedCell = row.createCell(4);
            notCompletedCell.setCellValue(stats.getNotCompletedCount());
            notCompletedCell.setCellStyle(dataStyle);

            // 6. % выполнения
            Cell percentCell = row.createCell(5);
            percentCell.setCellValue(stats.getCompletionPercentage() / 100.0);
            percentCell.setCellStyle(percentStyle);

            // 7. Распределение баллов - КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ

            // 7.1. Сначала создаём ВСЕ ячейки от G до P с одинаковым стилем
            for (int col = 6; col <= mergeEndColumn; col++) {
                Cell cell = row.createCell(col);
                cell.setCellStyle(distributionStyle);

                // Только в первую ячейку (G) добавляем значение
                if (col == 6) {
                    if (stats.getScoreDistribution() != null && !stats.getScoreDistribution().isEmpty()) {
                        String distribution = formatScoreDistribution(stats.getScoreDistribution());
                        cell.setCellValue(distribution);
                    } else {
                        cell.setCellValue("-");
                    }
                }
            }

            // 7.2. Теперь объединяем ячейки G-P
            sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 6, mergeEndColumn));
        }

        // ========== 5. НАСТРАИВАЕМ ШИРИНУ СТОЛБЦОВ ==========
        sheet.setColumnWidth(0, 8 * 256);    // № (уже включает "№" в значении)
        sheet.setColumnWidth(1, 12 * 256);   // Макс. балл
        sheet.setColumnWidth(2, 12 * 256);   // Полностью
        sheet.setColumnWidth(3, 12 * 256);   // Частично
        sheet.setColumnWidth(4, 15 * 256);   // Не справилось
        sheet.setColumnWidth(5, 15 * 256);   // % выполнения

        // Ширина для объединенной области G-P
        // Распределяем ширину равномерно на все объединенные столбцы
        for (int col = 6; col <= mergeEndColumn; col++) {
            sheet.setColumnWidth(col, 10 * 256);
        }

        return startRow + 1; // Пустая строка после блока
    }

    // ============ ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ============

    private String formatScoreDistribution(Map<Integer, Integer> distribution) {
        if (distribution == null || distribution.isEmpty()) {
            return "";
        }

        List<String> parts = new ArrayList<>();
        for (Map.Entry<Integer, Integer> entry : distribution.entrySet()) {
            int score = entry.getKey();  // количество баллов (0, 1, 2...)
            int count = entry.getValue(); // количество студентов

            String ballWord;
            if (score == 1) {
                ballWord = "балл";
            } else if (score >= 2 && score <= 4) {
                ballWord = "балла";
            } else {
                ballWord = "баллов";
            }

            parts.add(score + " " + ballWord + ": " + count);
        }
        return String.join("; ", parts);
    }

    // ============ СОХРАНЕНИЕ ФАЙЛА ============

    private File saveWorkbookToFile(Workbook workbook,
                                    TestSummaryDto testSummary) throws IOException {
        // Создаем директорию для отчетов, если её нет
        String safeSubject = testSummary.getSubject() != null ?
                testSummary.getSubject().replaceAll("[\\\\/:*?\"<>|]", "_"): "";
        String safeSchool = testSummary.getSchoolName() != null ?
                testSummary.getSchoolName().replaceAll("[\\\\/:*?\"<>|]", "_") : "";

        String folderPath = REPORTS_ANALISIS_BASE_FOLDER
                .replace("{школа}", safeSchool)
                .replace("{предмет}", safeSubject);

        if (!Files.exists(Path.of(folderPath))) {
            Files.createDirectories(Path.of(folderPath));
        }

        // Формируем имя файла
        String fileName = String.format("Детальный_отчет_%s_%s_%s.xlsx",
                testSummary.getSubject(),
                testSummary.getClassName(),
                testSummary.getTestDate().format(DateTimeFormatter.ofPattern("yyyyMMdd")));
        log.info("✅ название детального отчёта такое: {}", fileName);
        File file = Path.of(folderPath).resolve(fileName).toFile();

        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }

        log.info("✅ Детальный отчет сохранен: {}", file.getAbsolutePath());
        return file;
    }

    /**
     * Создает уникальное имя листа для рабочей книги
     */
    public String createUniqueSheetName(XSSFWorkbook workbook, TestSummaryDto testSummary) {
        String subjectCleaned = testSummary.getSubject().replaceAll("[^a-zA-Zа-яА-Я0-9_]", "");
        String subjectPart = subjectCleaned.substring(0, Math.min(12, subjectCleaned.length()));

        String classNameCleaned = testSummary.getClassName().replaceAll("[^a-zA-Zа-яА-Я0-9_]", "");

        String baseName = String.format("%s_%s_%s",
                subjectPart,
                classNameCleaned,
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

            // 2. Общая статистика теста (ОБНОВЛЕННАЯ ВЕРСИЯ)
            currentRow = createGeneralStatistics(workbook, sheet, testSummary, currentRow);

            // 3. Результаты студентов
            currentRow = createStudentsResults(workbook, sheet, studentResults, taskStatistics, currentRow);

            // 4. Анализ по заданиям (основная таблица)
            currentRow = createTaskAnalysis(workbook, sheet, taskStatistics, currentRow);

            // 5. Графики (если есть данные)
            if (taskStatistics != null && !taskStatistics.isEmpty() &&
                    studentResults != null && !studentResults.isEmpty()) {
                createGraphicalAnalysis(workbook, sheet, taskStatistics, currentRow);
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