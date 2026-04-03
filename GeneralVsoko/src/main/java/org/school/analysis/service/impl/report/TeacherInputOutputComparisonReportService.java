package org.school.analysis.service.impl.report;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.springframework.stereotype.Service;

import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Сервис построения сравнительного листа "Входная vs Выходная" для отчета учителя.
 */
@Service
@Slf4j
public class TeacherInputOutputComparisonReportService extends ExcelReportBase {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    public void addInputOutputComparisonSheets(XSSFWorkbook workbook,
                                               String teacherName,
                                               List<TeacherTestDetailDto> teacherTestDetails) {
        List<ComparisonPair> pairs = buildComparisonPairs(teacherTestDetails);

        if (pairs.isEmpty()) {
            log.info("Для учителя '{}' не найдены пары входной/выходной работ", teacherName);
            return;
        }

        int sheetIndex = 1;
        for (ComparisonPair pair : pairs) {
            String sheetName = buildUniqueSheetName(workbook, "Сравнение_" + sheetIndex++);
            XSSFSheet sheet = workbook.createSheet(sheetName);
            fillComparisonSheet(workbook, sheet, teacherName, pair);
        }
    }

    private void fillComparisonSheet(XSSFWorkbook workbook,
                                     XSSFSheet sheet,
                                     String teacherName,
                                     ComparisonPair pair) {
        int rowNum = 0;
        CellStyle titleStyle = getTitleStyle(workbook);
        CellStyle subtitleStyle = getSubtitleStyle(workbook);
        CellStyle headerStyle = getTableHeaderStyle(workbook);
        CellStyle percentStyle = getStyle(workbook, StyleType.PERCENT);
        CellStyle normalStyle = getStyle(workbook, StyleType.CENTERED);
        CellStyle decimalStyle = getStyle(workbook, StyleType.DECIMAL);

        createMergedTitle(sheet,
                "Сравнение входной и выходной работы",
                titleStyle, rowNum, 0, 4);
        rowNum++;

        Row infoRow = sheet.createRow(rowNum++);
        setCellValue(infoRow, 0, "Учитель", subtitleStyle);
        setCellValue(infoRow, 1, teacherName, subtitleStyle);
        setCellValue(infoRow, 2, "Предмет / класс", subtitleStyle);
        setCellValue(infoRow, 3, pair.output().getSubject() + " / " + pair.output().getClassName(), subtitleStyle);

        Row dateRow = sheet.createRow(rowNum++);
        setCellValue(dateRow, 0, "Входная", subtitleStyle);
        setCellValue(dateRow, 1, formatDate(pair.input().getTestDate()), subtitleStyle);
        setCellValue(dateRow, 2, "Выходная", subtitleStyle);
        setCellValue(dateRow, 3, formatDate(pair.output().getTestDate()), subtitleStyle);

        rowNum++;
        rowNum = createOverallComparisonBlock(sheet, rowNum, headerStyle, normalStyle, decimalStyle, percentStyle, pair);
        rowNum++;
        rowNum = createTaskSection(sheet, rowNum, "Результаты входной работы", headerStyle, percentStyle, pair.inputTaskPercentages());
        rowNum++;
        rowNum = createTaskSection(sheet, rowNum, "Результаты выходной работы", headerStyle, percentStyle, pair.outputTaskPercentages());
        rowNum++;
        int comparisonTableStart = rowNum;
        rowNum = createTaskDeltaSection(sheet, rowNum, headerStyle, percentStyle, pair.inputTaskPercentages(), pair.outputTaskPercentages());

        createComparisonLineChart(sheet, comparisonTableStart + 1, pair.taskNumbers().size(), rowNum + 1);
        autosize(sheet);
    }

    private int createOverallComparisonBlock(XSSFSheet sheet,
                                             int rowNum,
                                             CellStyle headerStyle,
                                             CellStyle normalStyle,
                                             CellStyle decimalStyle,
                                             CellStyle percentStyle,
                                             ComparisonPair pair) {
        Row header = sheet.createRow(rowNum++);
        setCellValue(header, 0, "Показатель", headerStyle);
        setCellValue(header, 1, "Входная", headerStyle);
        setCellValue(header, 2, "Выходная", headerStyle);
        setCellValue(header, 3, "Δ (выходная - входная)", headerStyle);

        rowNum = createMetricRow(sheet, rowNum, "Средний балл", pair.input().getAverageScore(), pair.output().getAverageScore(),
                normalStyle, decimalStyle);
        rowNum = createMetricRow(sheet, rowNum, "% выполнения теста",
                pair.input().getSuccessPercentage(), pair.output().getSuccessPercentage(),
                normalStyle, percentStyle);
        rowNum = createMetricRow(sheet, rowNum, "Присутствовало",
                toDouble(pair.input().getStudentsPresent()), toDouble(pair.output().getStudentsPresent()),
                normalStyle, normalStyle);
        return rowNum;
    }

    private int createMetricRow(XSSFSheet sheet,
                                int rowNum,
                                String metricName,
                                Double inputValue,
                                Double outputValue,
                                CellStyle titleStyle,
                                CellStyle valueStyle) {
        Row row = sheet.createRow(rowNum++);
        setCellValue(row, 0, metricName, titleStyle);
        setCellValue(row, 1, safe(inputValue), valueStyle);
        setCellValue(row, 2, safe(outputValue), valueStyle);
        setCellValue(row, 3, safe(outputValue) - safe(inputValue), valueStyle);
        return rowNum;
    }

    private int createTaskSection(XSSFSheet sheet,
                                  int rowNum,
                                  String title,
                                  CellStyle headerStyle,
                                  CellStyle percentStyle,
                                  Map<Integer, Double> taskValues) {
        Row titleRow = sheet.createRow(rowNum++);
        setCellValue(titleRow, 0, title, headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 2));

        Row header = sheet.createRow(rowNum++);
        setCellValue(header, 0, "Задание", headerStyle);
        setCellValue(header, 1, "% выполнения", headerStyle);
        setCellValue(header, 2, "Успех", headerStyle);

        for (Map.Entry<Integer, Double> entry : taskValues.entrySet()) {
            Row row = sheet.createRow(rowNum++);
            setCellValue(row, 0, entry.getKey(), getStyle(sheet.getWorkbook(), StyleType.CENTERED));
            setCellValue(row, 1, safe(entry.getValue()) / 100.0, percentStyle);
            setCellValue(row, 2, entry.getValue() >= 50.0 ? "✅" : "⚠️", getStyle(sheet.getWorkbook(), StyleType.CENTERED));
        }
        return rowNum;
    }

    private int createTaskDeltaSection(XSSFSheet sheet,
                                       int rowNum,
                                       CellStyle headerStyle,
                                       CellStyle percentStyle,
                                       Map<Integer, Double> inputValues,
                                       Map<Integer, Double> outputValues) {
        Row titleRow = sheet.createRow(rowNum++);
        setCellValue(titleRow, 0, "Сравнение входного и выходного теста по заданиям", headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 4));

        Row header = sheet.createRow(rowNum++);
        setCellValue(header, 0, "Задание", headerStyle);
        setCellValue(header, 1, "Входная", headerStyle);
        setCellValue(header, 2, "Выходная", headerStyle);
        setCellValue(header, 3, "Δ", headerStyle);
        setCellValue(header, 4, "Тренд", headerStyle);

        SortedSet<Integer> taskNumbers = new TreeSet<>();
        taskNumbers.addAll(inputValues.keySet());
        taskNumbers.addAll(outputValues.keySet());

        for (Integer taskNumber : taskNumbers) {
            double inputValue = safe(inputValues.get(taskNumber));
            double outputValue = safe(outputValues.get(taskNumber));
            double delta = outputValue - inputValue;

            Row row = sheet.createRow(rowNum++);
            setCellValue(row, 0, taskNumber, getStyle(sheet.getWorkbook(), StyleType.CENTERED));
            setCellValue(row, 1, inputValue / 100.0, percentStyle);
            setCellValue(row, 2, outputValue / 100.0, percentStyle);
            setCellValue(row, 3, delta / 100.0, percentStyle);
            setCellValue(row, 4, trend(delta), getStyle(sheet.getWorkbook(), StyleType.CENTERED));
        }
        return rowNum;
    }

    private void createComparisonLineChart(XSSFSheet sheet, int dataStartRow, int size, int chartStartRow) {
        if (size == 0) {
            return;
        }

        try {
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            var anchor = drawing.createAnchor(0, 0, 0, 0, 0, chartStartRow, 9, chartStartRow + 18);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Динамика по заданиям: входная vs выходная");

            XDDFCategoryAxis bottom = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottom.setTitle("Задание");
            XDDFValueAxis left = chart.createValueAxis(AxisPosition.LEFT);
            left.setTitle("% выполнения");
            left.setMinimum(0);
            left.setMaximum(1);
            left.setNumberFormat("0%");

            XDDFDataSource<Double> x = XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet, new CellRangeAddress(dataStartRow, dataStartRow + size - 1, 0, 0));
            XDDFNumericalDataSource<Double> input = XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet, new CellRangeAddress(dataStartRow, dataStartRow + size - 1, 1, 1));
            XDDFNumericalDataSource<Double> output = XDDFDataSourcesFactory.fromNumericCellRange(
                    sheet, new CellRangeAddress(dataStartRow, dataStartRow + size - 1, 2, 2));

            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottom, left);
            XDDFLineChartData.Series inputSeries = (XDDFLineChartData.Series) data.addSeries(x, input);
            inputSeries.setTitle("Входная", null);
            inputSeries.setSmooth(false);
            inputSeries.setMarkerStyle(MarkerStyle.CIRCLE);

            XDDFLineChartData.Series outputSeries = (XDDFLineChartData.Series) data.addSeries(x, output);
            outputSeries.setTitle("Выходная", null);
            outputSeries.setSmooth(false);
            outputSeries.setMarkerStyle(MarkerStyle.DIAMOND);

            chart.plot(data);
        } catch (Exception e) {
            log.error("Ошибка построения графика сравнения входной/выходной работы: {}", e.getMessage(), e);
        }
    }

    private List<ComparisonPair> buildComparisonPairs(List<TeacherTestDetailDto> details) {
        Map<String, List<TeacherTestDetailDto>> grouped = details.stream()
                .filter(Objects::nonNull)
                .filter(detail -> detail.getTestSummary() != null)
                .collect(Collectors.groupingBy(detail -> detail.getTestSummary().getSubject() + "||"
                        + detail.getTestSummary().getClassName()));

        List<ComparisonPair> pairs = new ArrayList<>();
        for (List<TeacherTestDetailDto> group : grouped.values()) {
            Optional<TeacherTestDetailDto> input = group.stream()
                    .filter(this::isInputWork)
                    .max(Comparator.comparing(d -> d.getTestSummary().getTestDate()));

            Optional<TeacherTestDetailDto> output = group.stream()
                    .filter(this::isOutputWork)
                    .max(Comparator.comparing(d -> d.getTestSummary().getTestDate()));

            if (input.isPresent() && output.isPresent()) {
                pairs.add(buildPair(input.get(), output.get()));
            }
        }
        return pairs;
    }

    private ComparisonPair buildPair(TeacherTestDetailDto input, TeacherTestDetailDto output) {
        SortedSet<Integer> taskNumbers = new TreeSet<>();
        taskNumbers.addAll(readTaskPercentages(input.getTaskStatistics()).keySet());
        taskNumbers.addAll(readTaskPercentages(output.getTaskStatistics()).keySet());

        Map<Integer, Double> inputValues = new LinkedHashMap<>();
        Map<Integer, Double> outputValues = new LinkedHashMap<>();
        Map<Integer, Double> rawInput = readTaskPercentages(input.getTaskStatistics());
        Map<Integer, Double> rawOutput = readTaskPercentages(output.getTaskStatistics());

        for (Integer taskNumber : taskNumbers) {
            inputValues.put(taskNumber, safe(rawInput.get(taskNumber)));
            outputValues.put(taskNumber, safe(rawOutput.get(taskNumber)));
        }

        return new ComparisonPair(input.getTestSummary(), output.getTestSummary(), taskNumbers, inputValues, outputValues);
    }

    private Map<Integer, Double> readTaskPercentages(Map<Integer, TaskStatisticsDto> taskStatistics) {
        if (taskStatistics == null || taskStatistics.isEmpty()) {
            return Map.of();
        }
        return taskStatistics.entrySet().stream()
                .collect(Collectors.toMap(
                        Map.Entry::getKey,
                        entry -> entry.getValue().getCompletionPercentage(),
                        (a, b) -> a,
                        TreeMap::new
                ));
    }

    private boolean isInputWork(TeacherTestDetailDto dto) {
        return hasTypeToken(dto, "вход");
    }

    private boolean isOutputWork(TeacherTestDetailDto dto) {
        return hasTypeToken(dto, "выход");
    }

    private boolean hasTypeToken(TeacherTestDetailDto dto, String token) {
        String testType = Optional.ofNullable(dto.getTestSummary())
                .map(TestSummaryDto::getTestType)
                .orElse("");
        return testType.toLowerCase(Locale.ROOT).contains(token);
    }

    private String buildUniqueSheetName(XSSFWorkbook workbook, String base) {
        String candidate = base;
        int idx = 1;
        while (workbook.getSheet(candidate) != null) {
            candidate = base + "_" + idx++;
        }
        return candidate;
    }

    private String trend(double delta) {
        if (delta > 0.01) {
            return "Рост";
        }
        if (delta < -0.01) {
            return "Снижение";
        }
        return "Без изменений";
    }

    private Double toDouble(Integer value) {
        return value == null ? 0.0 : value.doubleValue();
    }

    private double safe(Double value) {
        return value == null ? 0.0 : value;
    }

    private String formatDate(java.time.LocalDate date) {
        return date == null ? "-" : date.format(DATE_FORMATTER);
    }

    private void autosize(Sheet sheet) {
        for (int i = 0; i <= 8; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private record ComparisonPair(
            TestSummaryDto input,
            TestSummaryDto output,
            SortedSet<Integer> taskNumbers,
            Map<Integer, Double> inputTaskPercentages,
            Map<Integer, Double> outputTaskPercentages
    ) {
    }
}
