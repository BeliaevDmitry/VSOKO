package org.school.analysis.service.impl.report;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.util.DateTimeFormatters;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

@Component
@Slf4j
public class TeacherReportGenerator extends ExcelReportBase {

    private final DetailReportGenerator detailReportGenerator;

    // Предопределенные ширины колонок в символах (уже с учетом фильтров)
    private static final int[] TEACHER_SUMMARY_WIDTHS = {
            19, // 0: Предмет
            12, // 1: Класс
            12, // 2: Дата (дд.мм.гггг)
            15, // 3: Тип теста
            17, // 4: Присутствовало
            17, // 5: Отсутствовало
            10,  // 6: Всего
            15, // 7: % присутствия
            17, // 8: Средний балл
            18  // 9: % выполнения
    };

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
                createTeacherComparativeSheets(workbook, teacherTests, teacherTestDetails);

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

        // 1. ЗАГОЛОВОК ОТЧЕТА
        createReportHeader(sheet, workbook, teacherName, teacherTests);

        // 2. ЗАГОЛОВКИ ТАБЛИЦЫ
        Row headerRow = createTableHeaders(sheet, workbook);

        // 3. ДАННЫЕ ТЕСТОВ
        fillTeacherTestData(sheet, workbook, teacherTests, headerRow.getRowNum() + 1);

        // 4. ИТОГОВАЯ СТРОКА
        if (!teacherTests.isEmpty()) {
            addTeacherSummaryRow(sheet, workbook, teacherTests);
        }

        // 5. НАСТРОЙКА ФУНКЦИОНАЛЬНОСТИ
        setupTableFeatures(sheet, headerRow, TEACHER_SUMMARY_WIDTHS.length);
    }

    /**
     * Создает заголовок отчета и информационную строку
     */
    private void createReportHeader(Sheet sheet, Workbook workbook,
                                    String teacherName, List<TestSummaryDto> teacherTests) {

        CellStyle titleStyle = getTitleStyle(workbook);
        CellStyle infoStyle = getSubtitleStyle(workbook);

        // Заголовок отчета
        createMergedTitle(sheet,
                "Отчет по тестам учителя: " + teacherName,
                titleStyle,
                0, 0, TEACHER_SUMMARY_WIDTHS.length - 1);

        // Информационная строка
        Row infoRow = sheet.createRow(1);

        // Дата формирования (левая часть)
        Cell dateCell = infoRow.createCell(0);
        dateCell.setCellValue("Отчет сгенерирован: " +
                LocalDateTime.now().format(DateTimeFormatters.DISPLAY_DATE));
        dateCell.setCellStyle(infoStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 4));

        // Количество тестов (правая часть)
        Cell testCountCell = infoRow.createCell(5);
        testCountCell.setCellValue("Тестов: " + teacherTests.size());
        testCountCell.setCellStyle(infoStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 5, TEACHER_SUMMARY_WIDTHS.length - 1));

        // Пустая строка для разделения
        sheet.createRow(2);
    }

    /**
     * Создает заголовки таблицы
     */
    private Row createTableHeaders(Sheet sheet, Workbook workbook) {
        String[] headers = {
                "Предмет", "Класс", "Дата", "Тип",
                "Присутствовало", "Отсутствовало", "Всего", "% присутствия",
                "Средний балл", "% выполнения"
        };

        Row headerRow = sheet.createRow(3);
        CellStyle tableHeaderStyle = getTableHeaderStyle(workbook);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(tableHeaderStyle);
        }

        return headerRow;
    }

    /**
     * Заполняет данные тестов с правильным форматированием
     */
    private void fillTeacherTestData(Sheet sheet, Workbook workbook,
                                     List<TestSummaryDto> teacherTests, int startRow) {

        // Получаем стили из кэша
        CellStyle normalStyle = getStyle(workbook, StyleType.NORMAL);
        CellStyle percentStyle = getStyle(workbook, StyleType.PERCENT);
        CellStyle decimalStyle = getStyle(workbook, StyleType.DECIMAL);
        CellStyle centeredStyle = getStyle(workbook, StyleType.CENTERED);

        // Маппинг стилей по колонкам
        CellStyle[] columnStyles = {
                normalStyle,    // 0: Предмет
                normalStyle,    // 1: Класс
                centeredStyle,  // 2: Дата
                centeredStyle,  // 3: Тип
                centeredStyle,  // 4: Присутствовало
                centeredStyle,  // 5: Отсутствовало
                centeredStyle,  // 6: Всего
                percentStyle,   // 7: % присутствия
                decimalStyle,   // 8: Средний балл
                percentStyle    // 9: % выполнения
        };

        int rowNum = startRow;
        for (TestSummaryDto test : teacherTests) {
            Row row = sheet.createRow(rowNum++);
            fillTeacherTestRow(row, test, columnStyles);
        }
    }

    /**
     * Заполняет одну строку с данными теста
     */
    private void fillTeacherTestRow(Row row, TestSummaryDto test, CellStyle[] columnStyles) {
        // Текстовые данные
        setCellValue(row, 0, test.getSubject(), columnStyles[0]);
        setCellValue(row, 1, test.getClassName(), columnStyles[1]);
        setCellValue(row, 2,
                test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE),
                columnStyles[2]);
        setCellValue(row, 3, test.getTestType(), columnStyles[3]);

        // Числовые данные
        setCellValue(row, 4, test.getStudentsPresent(), columnStyles[4]);
        setCellValue(row, 5, test.getStudentsAbsent(), columnStyles[5]);
        setCellValue(row, 6, test.getClassSize(), columnStyles[6]);

        // Проценты
        double attendancePercent = test.getAttendancePercentage() != null ?
                test.getAttendancePercentage() / 100.0 : 0.0;
        setCellValue(row, 7, attendancePercent, columnStyles[7]);

        // Десятичные числа
        double avgScore = test.getAverageScore() != null ? test.getAverageScore() : 0.0;
        setCellValue(row, 8, avgScore, columnStyles[8]);

        // Проценты выполнения
        double successPercent = test.getSuccessPercentage() != null ?
                test.getSuccessPercentage() / 100.0 : 0.0;
        setCellValue(row, 9, successPercent, columnStyles[9]);
    }

    /**
     * Добавляет итоговую строку со средними показателями
     */
    private void addTeacherSummaryRow(Sheet sheet, Workbook workbook,
                                      List<TestSummaryDto> tests) {

        int lastRow = sheet.getLastRowNum() + 2; // Пропускаем строку
        Row summaryRow = sheet.createRow(lastRow);

        // Получаем стили
        CellStyle summaryStyle = createSummaryStyle(workbook);
        CellStyle percentStyle = getStyle(workbook, StyleType.PERCENT);
        CellStyle decimalStyle = getStyle(workbook, StyleType.DECIMAL);

        // Заголовок итогов
        summaryRow.createCell(0).setCellValue("Средние показатели:");
        summaryRow.getCell(0).setCellStyle(summaryStyle);

        // Рассчитываем средние значения
        double[] averages = calculateAverages(tests);

        // Заполняем ячейки
        setSummaryCell(summaryRow, 4, averages[0], summaryStyle); // Присутствовало
        setSummaryCell(summaryRow, 5, averages[1], summaryStyle); // Отсутствовало
        setSummaryCell(summaryRow, 6, averages[2], summaryStyle); // Всего
        setSummaryCell(summaryRow, 7, averages[3], percentStyle); // % присутствия
        setSummaryCell(summaryRow, 8, averages[4], decimalStyle); // Средний балл
        setSummaryCell(summaryRow, 9, averages[5], percentStyle); // % выполнения
    }

    /**
     * Рассчитывает средние значения по всем тестам
     */
    private double[] calculateAverages(List<TestSummaryDto> tests) {
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

        return new double[]{avgPresent, avgAbsent, avgClassSize,
                avgAttendance, avgScore, avgSuccess};
    }

    /**
     * Устанавливает значение в ячейку итоговой строки
     */
    private void setSummaryCell(Row row, int column, double value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    /**
     * Настраивает функциональность таблицы (фильтры, закрепление, ширины)
     */
    private void setupTableFeatures(Sheet sheet, Row headerRow, int columnCount) {
        try {
            // 1. Устанавливаем ширины колонок (ПЕРВЫМ делом!)
            setColumnWidthsWithFilterMargin(sheet, columnCount);

            // 2. Включаем автофильтр
            enableAutoFilter(sheet, headerRow, columnCount);

            // 3. Закрепляем первые 4 строки
            freezeFirstNRows(sheet, 4);

            // 4. Настраиваем высоту строки заголовков (для лучшего отображения фильтров)
            headerRow.setHeight((short) (sheet.getDefaultRowHeight() * 1.2));

            log.info("✅ Настройки листа 'Сводка по тестам' применены");

        } catch (Exception e) {
            log.error("❌ Ошибка при настройке листа: {}", e.getMessage(), e);
        }
    }

    /**
     * Устанавливает фиксированные ширины колонок с учетом фильтров
     */
    private void setColumnWidthsWithFilterMargin(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount && i < TEACHER_SUMMARY_WIDTHS.length; i++) {
            int widthInTwips = TEACHER_SUMMARY_WIDTHS[i] * 256;
            sheet.setColumnWidth(i, widthInTwips);
            log.debug("Установлена ширина колонки {}: {} символов",
                    i, TEACHER_SUMMARY_WIDTHS[i]);
        }
    }

    /**
     * Включает автофильтр на заголовках таблицы
     */
    private void enableAutoFilter(Sheet sheet, Row headerRow, int columnCount) {
        if (sheet.getLastRowNum() > headerRow.getRowNum()) {
            CellRangeAddress filterRange = new CellRangeAddress(
                    headerRow.getRowNum(),
                    sheet.getLastRowNum(),
                    0,
                    columnCount - 1
            );

            sheet.setAutoFilter(filterRange);
            log.debug("Автофильтр установлен на диапазоне A{}-{}{}",
                    headerRow.getRowNum() + 1, // Excel 1-based
                    getExcelColumnName(columnCount - 1),
                    sheet.getLastRowNum() + 1);
        }
    }

    /**
     * Закрепляет первые N строк листа
     */
    private void freezeFirstNRows(Sheet sheet, int rowsCount) {
        try {
            sheet.createFreezePane(0, rowsCount);
            log.debug("Закреплены первые {} строк", rowsCount);
        } catch (Exception e) {
            log.error("Ошибка при закреплении строк: {}", e.getMessage());
        }
    }

    /**
     * Создает именованный диапазон для удобства
     */
    private void createNamedRange(Sheet sheet, Row headerRow, int columnCount) {
        try {
            Name namedRange = sheet.getWorkbook().createName();
            namedRange.setNameName("TeacherTestData");
            namedRange.setRefersToFormula(
                    String.format("'%s'!$A$%d:$%s$%d",
                            sheet.getSheetName(),
                            headerRow.getRowNum() + 1,
                            getExcelColumnName(columnCount - 1),
                            sheet.getLastRowNum() + 1)
            );
        } catch (Exception e) {
            log.warn("Не удалось создать именованный диапазон: {}", e.getMessage());
        }
    }

    private void createTeacherTestDetailSheet(XSSFWorkbook workbook,
                                              TeacherTestDetailDto testDetail) {

        TestSummaryDto testSummary = testDetail.getTestSummary();
        if (testSummary == null) {
            log.warn("Пропускаем тест без основных данных");
            return;
        }

        String sheetName = detailReportGenerator.createUniqueSheetName(workbook, testSummary);

        try {
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

    private void createTeacherComparativeSheets(XSSFWorkbook workbook,
                                                List<TestSummaryDto> teacherTests,
                                                List<TeacherTestDetailDto> teacherTestDetails) {
        Map<String, Map<Integer, TaskStatisticsDto>> statsById = teacherTestDetails.stream()
                .filter(d -> d.getTestSummary() != null && d.getTestSummary().getReportFileId() != null)
                .collect(Collectors.toMap(
                        d -> d.getTestSummary().getReportFileId(),
                        TeacherTestDetailDto::getTaskStatistics,
                        (first, second) -> first
                ));

        Map<String, List<TestSummaryDto>> bySubjectClass = teacherTests.stream()
                .filter(t -> t.getTestType() != null && t.getTestType().startsWith("ЕГКР"))
                .collect(Collectors.groupingBy(t -> t.getSubject() + "||" + t.getClassName()));

        for (Map.Entry<String, List<TestSummaryDto>> entry : bySubjectClass.entrySet()) {
            List<TestSummaryDto> egkrTests = entry.getValue().stream()
                    .sorted(Comparator.comparing(TestSummaryDto::getTestDate))
                    .collect(Collectors.toList());

            if (egkrTests.size() < 2) {
                continue;
            }

            List<TestSummaryDto> selected = new ArrayList<>(egkrTests.subList(egkrTests.size() - 2, egkrTests.size()));
            Optional<TestSummaryDto> ege = teacherTests.stream()
                    .filter(t -> Objects.equals(t.getSubject(), selected.get(0).getSubject()))
                    .filter(t -> Objects.equals(t.getClassName(), selected.get(0).getClassName()))
                    .filter(t -> t.getTestType() != null && t.getTestType().contains("ЕГЭ"))
                    .max(Comparator.comparing(TestSummaryDto::getTestDate));
            ege.ifPresent(selected::add);

            String sheetName = createSafeUniqueSheetName(
                    workbook, "Сравнение_" + selected.get(0).getSubject() + "_" + selected.get(0).getClassName());
            createTeacherComparisonSheet(workbook, sheetName, selected, statsById);
        }
    }

    private void createTeacherComparisonSheet(XSSFWorkbook workbook,
                                              String sheetName,
                                              List<TestSummaryDto> tests,
                                              Map<String, Map<Integer, TaskStatisticsDto>> statsById) {
        Sheet sheet = workbook.createSheet(sheetName);
        Row title = sheet.createRow(0);
        title.createCell(0).setCellValue("Сравнительный анализ");

        Row header = sheet.createRow(2);
        header.createCell(0).setCellValue("Дата");
        header.createCell(1).setCellValue("Тип");
        header.createCell(2).setCellValue("% выполнения");
        header.createCell(3).setCellValue("Средний балл");
        header.createCell(4).setCellValue("% присутствия");
        header.createCell(5).setCellValue("Участников");

        int row = 3;
        for (TestSummaryDto test : tests) {
            Row data = sheet.createRow(row++);
            data.createCell(0).setCellValue(test.getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
            data.createCell(1).setCellValue(test.getTestType());
            data.createCell(2).setCellValue(test.getSuccessPercentage());
            data.createCell(3).setCellValue(test.getAverageScore() != null ? test.getAverageScore() : 0.0);
            data.createCell(4).setCellValue(test.getAttendancePercentage());
            data.createCell(5).setCellValue(test.getStudentsPresent() != null ? test.getStudentsPresent() : 0);
        }

        int taskHeaderRowNum = row + 2;
        Row taskHeader = sheet.createRow(taskHeaderRowNum);
        taskHeader.createCell(0).setCellValue("Задание");
        for (int i = 0; i < tests.size(); i++) {
            taskHeader.createCell(i + 1).setCellValue(
                    tests.get(i).getTestType() + " " + tests.get(i).getTestDate().format(DateTimeFormatters.DISPLAY_DATE));
        }

        Map<Integer, List<Double>> taskData = buildTaskData(tests, statsById);
        int taskRowNum = taskHeaderRowNum + 1;
        for (Map.Entry<Integer, List<Double>> entry : taskData.entrySet()) {
            Row taskRow = sheet.createRow(taskRowNum++);
            taskRow.createCell(0).setCellValue(entry.getKey());
            for (int i = 0; i < entry.getValue().size(); i++) {
                taskRow.createCell(i + 1).setCellValue(entry.getValue().get(i));
            }
        }

        if (!taskData.isEmpty()) {
            createTaskComparisonLineChart(sheet, taskHeaderRowNum, taskRowNum - 1, tests.size());
        }
    }

    private Map<Integer, List<Double>> buildTaskData(List<TestSummaryDto> tests,
                                                     Map<String, Map<Integer, TaskStatisticsDto>> statsById) {
        List<Map<Integer, TaskStatisticsDto>> stats = tests.stream()
                .map(t -> statsById.getOrDefault(t.getReportFileId(), Map.of()))
                .collect(Collectors.toList());

        Set<Integer> allTasks = new TreeSet<>();
        stats.forEach(map -> allTasks.addAll(map.keySet()));

        Map<Integer, List<Double>> result = new LinkedHashMap<>();
        for (Integer task : allTasks) {
            List<Double> values = new ArrayList<>();
            for (Map<Integer, TaskStatisticsDto> map : stats) {
                TaskStatisticsDto dto = map.get(task);
                values.add(dto != null ? dto.getCompletionPercentage() : 0.0);
            }
            result.put(task, values);
        }
        return result;
    }

    private void createTaskComparisonLineChart(Sheet sheet, int headerRow, int dataEndRow, int seriesCount) {
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 0, dataEndRow + 2, 10, dataEndRow + 18));
        chart.setTitleText("Сравнение выполнения заданий");

        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        xAxis.setTitle("Задание");
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle("% выполнения");

        XDDFDataSource<Double> categories = XDDFDataSourcesFactory.fromNumericCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new CellRangeAddress(headerRow + 1, dataEndRow, 0, 0)
        );
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, xAxis, yAxis);

        for (int i = 1; i <= seriesCount; i++) {
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                    (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                    new CellRangeAddress(headerRow + 1, dataEndRow, i, i)
            );
            XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(categories, values);
            series.setTitle(sheet.getRow(headerRow).getCell(i).getStringCellValue(), null);
            series.setMarkerStyle(MarkerStyle.CIRCLE);
        }

        chart.plot(data);
    }

    private String createSafeUniqueSheetName(XSSFWorkbook workbook, String baseName) {
        String sanitized = baseName.replaceAll("[\\\\/*\\[\\]:?]", "_");
        sanitized = sanitized.length() > 25 ? sanitized.substring(0, 25) : sanitized;
        String candidate = sanitized;
        int suffix = 1;
        while (workbook.getSheet(candidate) != null) {
            String tail = "_" + suffix++;
            int maxBase = Math.max(1, 31 - tail.length());
            candidate = sanitized.substring(0, Math.min(sanitized.length(), maxBase)) + tail;
        }
        return candidate;
    }

    /**
     * Преобразует индекс колонки в буквенное обозначение Excel (A, B, C, ... AA, AB, ...)
     */
    private String getExcelColumnName(int columnIndex) {
        StringBuilder columnName = new StringBuilder();

        while (columnIndex >= 0) {
            int remainder = columnIndex % 26;
            columnName.insert(0, (char) ('A' + remainder));
            columnIndex = (columnIndex / 26) - 1;
        }

        return columnName.toString();
    }
}
