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

        // Заголовок раздела
        Row sectionHeader = sheet.createRow(startRow++);
        Cell headerCell = sectionHeader.createCell(0);
        headerCell.setCellValue("ГРАФИЧЕСКИЙ АНАЛИЗ");
        headerCell.setCellStyle(createSectionHeaderStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));

        // Описание - откуда берутся данные
        Row descriptionRow = sheet.createRow(startRow++);
        Cell descCell = descriptionRow.createCell(0);
        descCell.setCellValue("*Данные для графиков взяты из раздела 'АНАЛИЗ ПО ЗАДАНИЯМ'");
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));

        // Стиль для описания
        CellStyle descStyle = workbook.createCellStyle();
        Font descFont = workbook.createFont();
        descFont.setItalic(true);
        descFont.setFontHeightInPoints((short) 10);
        descFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        descStyle.setFont(descFont);
        descCell.setCellStyle(descStyle);

        startRow++; // Пустая строка

        // Если нет данных для графиков - выводим сообщение
        if (taskStatistics == null || taskStatistics.isEmpty()) {
            Row noDataRow = sheet.createRow(startRow++);
            noDataRow.createCell(0).setCellValue("Нет данных для графиков");
            return startRow;
        }

        // Проходим по всем строкам и ищем заголовок "АНАЛИЗ ПО ЗАДАНИЯМ"
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

        // Если таблица найдена - создаем графики на ее основе
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
            return sheet.getLastRowNum() + 2;
        }
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
     * Создает общую статистику теста (ОБНОВЛЕННАЯ ВЕРСИЯ)
     */
    private int createGeneralStatistics(Workbook workbook, Sheet sheet,
                                        TestSummaryDto testSummary, int startRow) {

        // =================== ЗАГОЛОВОК РАЗДЕЛА ===================
        Row headerRow = sheet.createRow(startRow++);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("ОБЩАЯ СТАТИСТИКА ТЕСТА");

        // Стиль для заголовка раздела БЕЗ границ
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints(SECTION_FONT_SIZE);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(SECTION_HEADER_BG_COLOR);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // УБИРАЕМ ВСЕ ГРАНИЦЫ
        headerStyle.setBorderTop(BorderStyle.NONE);
        headerStyle.setBorderBottom(BorderStyle.NONE);
        headerStyle.setBorderLeft(BorderStyle.NONE);
        headerStyle.setBorderRight(BorderStyle.NONE);
        headerStyle.setAlignment(HorizontalAlignment.LEFT);
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, 3));

        // =================== ДАННЫЕ СТАТИСТИКИ ===================
        // Создаем стиль для данных статистики С ГРАНИЦАМИ
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setWrapText(true);

        // Стиль для меток (левая колонка)
        CellStyle labelStyle = workbook.createCellStyle();
        labelStyle.cloneStyleFrom(dataStyle);
        labelStyle.setAlignment(HorizontalAlignment.RIGHT);
        labelStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        labelStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font labelFont = workbook.createFont();
        labelFont.setBold(true);
        labelStyle.setFont(labelFont);

        // Стиль для значений (правая колонка)
        CellStyle valueStyle = workbook.createCellStyle();
        valueStyle.cloneStyleFrom(dataStyle);
        valueStyle.setAlignment(HorizontalAlignment.CENTER);

        // 1. Первая группа данных (посещаемость)
        Row row1 = sheet.createRow(startRow++);

        // Всего учеников в классе
        Cell label1 = row1.createCell(1);
        label1.setCellValue("Всего учеников в классе");
        label1.setCellStyle(labelStyle);

        Cell value1 = row1.createCell(3); // Перемещаем в колонку D
        Integer totalStudents = testSummary.getClassSize() != null ?
                testSummary.getClassSize() : testSummary.getStudentsTotal();
        value1.setCellValue(totalStudents != null ? totalStudents : 0);
        value1.setCellStyle(valueStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2)); // Объединяем B-C

        // Присутствовало на тесте
        Row row2 = sheet.createRow(startRow++);
        Cell label2 = row2.createCell(1);
        label2.setCellValue("Присутствовало на тесте");
        label2.setCellStyle(labelStyle);

        Cell value2 = row2.createCell(3);
        value2.setCellValue(testSummary.getStudentsPresent() != null ? testSummary.getStudentsPresent() : 0);
        value2.setCellStyle(valueStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        // Отсутствовало на тесте
        Row row3 = sheet.createRow(startRow++);
        Cell label3 = row3.createCell(1);
        label3.setCellValue("Отсутствовало на тесте");
        label3.setCellStyle(labelStyle);

        Cell value3 = row3.createCell(3);
        value3.setCellValue(testSummary.getStudentsAbsent() != null ? testSummary.getStudentsAbsent() : 0);
        value3.setCellStyle(valueStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        // Процент присутствия
        Row row4 = sheet.createRow(startRow++);
        Cell label4 = row4.createCell(1);
        label4.setCellValue("Процент присутствия");
        label4.setCellStyle(labelStyle);

        Cell value4 = row4.createCell(3);
        double attendancePercentage = 0.0;
        if (testSummary.getStudentsPresent() != null && testSummary.getClassSize() != null &&
                testSummary.getClassSize() > 0) {
            attendancePercentage = (testSummary.getStudentsPresent() * 100.0) / testSummary.getClassSize();
        }
        value4.setCellValue(attendancePercentage / 100.0);
        value4.setCellStyle(createPercentStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        // 2. Вторая группа данных (результаты)
        Row row5 = sheet.createRow(startRow++);
        Cell label5 = row5.createCell(1);
        label5.setCellValue("Количество заданий");
        label5.setCellStyle(labelStyle);

        Cell value5 = row5.createCell(3);
        value5.setCellValue(testSummary.getTaskCount() != null ? testSummary.getTaskCount() : 0);
        value5.setCellStyle(valueStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        Row row6 = sheet.createRow(startRow++);
        Cell label6 = row6.createCell(1);
        label6.setCellValue("Максимальный балл за тест");
        label6.setCellStyle(labelStyle);

        Cell value6 = row6.createCell(3);
        value6.setCellValue(testSummary.getMaxTotalScore() != null ? testSummary.getMaxTotalScore() : 0);
        value6.setCellStyle(valueStyle);
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        Row row7 = sheet.createRow(startRow++);
        Cell label7 = row7.createCell(1);
        label7.setCellValue("Средний балл по классу");
        label7.setCellStyle(labelStyle);

        Cell value7 = row7.createCell(3);
        Double averageScore = testSummary.getAverageScore();
        value7.setCellValue(averageScore != null ? averageScore : 0.0);
        value7.setCellStyle(createDecimalStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        Row row8 = sheet.createRow(startRow++);
        Cell label8 = row8.createCell(1);
        label8.setCellValue("Процент выполнения теста");
        label8.setCellStyle(labelStyle);

        Cell value8 = row8.createCell(3);
        double successPercentage = 0.0;
        if (testSummary.getMaxTotalScore() != null && testSummary.getMaxTotalScore() > 0 &&
                averageScore != null) {
            successPercentage = (averageScore * 100.0) / testSummary.getMaxTotalScore();
        }
        value8.setCellValue(successPercentage / 100.0);
        value8.setCellStyle(createPercentStyle(workbook));
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 1, 2));

        // Устанавливаем ширину колонок для красивого отображения
        sheet.setColumnWidth(1, 20 * 256); // Колонка B
        sheet.setColumnWidth(2, 1 * 256);  // Колонка C (очень узкая, почти не видна)
        sheet.setColumnWidth(3, 12 * 256); // Колонка D

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
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));

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
            headers[5 + i] = Integer.toString(i);
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
        sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, HEADER_MERGE_COUNT_TEST));

        // Заголовки таблицы
        Row columnHeaderRow = sheet.createRow(startRow++);
        String[] headers = {
                "№", "Макс. балл", "Полностью", "Частично", "Не справилось",
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