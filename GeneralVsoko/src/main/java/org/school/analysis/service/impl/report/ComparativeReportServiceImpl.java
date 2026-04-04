package org.school.analysis.service.impl.report;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.AnalysisService;
import org.school.analysis.service.ComparativeReportService;
import org.springframework.stereotype.Service;

import java.io.File;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class ComparativeReportServiceImpl extends ExcelReportBase implements ComparativeReportService {

    private final AnalysisService analysisService;

    @Override
    public File generateEgkrEgeComparativeReport(String school, String currentAcademicYear) {
        try {
            List<TestSummaryDto> allTests = analysisService.getAllTestsSummary(school, currentAcademicYear);
            List<ComparisonGroup> groups = buildGroups(allTests);

            if (groups.isEmpty()) {
                log.warn("Нет данных для сравнительного отчета ЕГКР/ЕГЭ");
                return null;
            }

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Map<ComparisonGroup, String> sheetNames = buildUniqueSheetNames(groups);
                createSummarySheet(workbook, sheetNames);

                for (Map.Entry<ComparisonGroup, String> entry : sheetNames.entrySet()) {
                    createComparisonSheet(workbook, entry.getKey(), entry.getValue());
                }

                return saveWorkbook(
                        workbook,
                        createReportsFolder(school),
                        "ЕГКР_ЕГЭ_сравнение.xlsx"
                );
            }
        } catch (Exception e) {
            log.error("Ошибка генерации сравнительного отчета ЕГКР/ЕГЭ", e);
            return null;
        }
    }

    @Override
    public File generateEgkrEgeSubjectComparativeReport(String school, String currentAcademicYear) {
        try {
            List<TestSummaryDto> allTests = analysisService.getAllTestsSummary(school, currentAcademicYear);
            List<SubjectComparisonGroup> subjectGroups = buildSubjectGroups(allTests);

            if (subjectGroups.isEmpty()) {
                log.warn("Нет данных для сравнительного отчета ЕГКР/ЕГЭ по предметам");
                return null;
            }

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Map<SubjectComparisonGroup, String> sheetNames = buildUniqueSubjectSheetNames(subjectGroups);
                createSubjectSummarySheet(workbook, sheetNames);
                for (Map.Entry<SubjectComparisonGroup, String> entry : sheetNames.entrySet()) {
                    createSubjectComparisonSheet(workbook, entry.getKey(), entry.getValue());
                }

                return saveWorkbook(
                        workbook,
                        createReportsFolder(school),
                        "ЕГКР_ЕГЭ_сравнение_по_предметам.xlsx"
                );
            }
        } catch (Exception e) {
            log.error("Ошибка генерации сравнительного отчета ЕГКР/ЕГЭ по предметам", e);
            return null;
        }
    }

    private Map<SubjectComparisonGroup, String> buildUniqueSubjectSheetNames(List<SubjectComparisonGroup> groups) {
        Map<SubjectComparisonGroup, String> names = new LinkedHashMap<>();
        Set<String> usedNames = new HashSet<>();
        for (SubjectComparisonGroup group : groups) {
            String base = group.sheetName();
            String candidate = base;
            int suffix = 2;
            while (usedNames.contains(candidate)) {
                String suffixPart = "_" + suffix++;
                int maxBaseLength = Math.max(1, 31 - suffixPart.length());
                String truncatedBase = base.length() > maxBaseLength ? base.substring(0, maxBaseLength) : base;
                candidate = truncatedBase + suffixPart;
            }
            usedNames.add(candidate);
            names.put(group, candidate);
        }
        return names;
    }

    private List<SubjectComparisonGroup> buildSubjectGroups(List<TestSummaryDto> allTests) {
        Map<String, List<TestSummaryDto>> bySubject = allTests.stream()
                .filter(t -> t.getSubject() != null && !t.getSubject().isBlank())
                .filter(t -> t.getClassName() != null && !t.getClassName().isBlank())
                .collect(Collectors.groupingBy(TestSummaryDto::getSubject));

        List<SubjectComparisonGroup> result = new ArrayList<>();
        for (Map.Entry<String, List<TestSummaryDto>> entry : bySubject.entrySet()) {
            Map<String, List<TestSummaryDto>> byClass = entry.getValue().stream()
                    .collect(Collectors.groupingBy(TestSummaryDto::getClassName));

            List<ClassComparison> classComparisons = new ArrayList<>();
            for (Map.Entry<String, List<TestSummaryDto>> classEntry : byClass.entrySet()) {
                List<TestSummaryDto> egkr = classEntry.getValue().stream()
                        .filter(t -> t.getTestType() != null && t.getTestType().startsWith("ЕГКР"))
                        .sorted(Comparator.comparing(TestSummaryDto::getTestDate))
                        .collect(Collectors.toList());
                if (egkr.size() < 2) {
                    continue;
                }

                TestSummaryDto egkr1 = egkr.get(egkr.size() - 2);
                TestSummaryDto egkr2 = egkr.get(egkr.size() - 1);
                Optional<TestSummaryDto> ege = classEntry.getValue().stream()
                        .filter(t -> t.getTestType() != null && t.getTestType().contains("ЕГЭ"))
                        .max(Comparator.comparing(TestSummaryDto::getTestDate));

                classComparisons.add(new ClassComparison(classEntry.getKey(), egkr1, egkr2, ege.orElse(null)));
            }

            if (!classComparisons.isEmpty()) {
                classComparisons.sort(Comparator.comparing(ClassComparison::className));
                result.add(new SubjectComparisonGroup(entry.getKey(), classComparisons));
            }
        }

        result.sort(Comparator.comparing(SubjectComparisonGroup::subject));
        return result;
    }

    private void createSubjectSummarySheet(XSSFWorkbook workbook, Map<SubjectComparisonGroup, String> groups) {
        Sheet sheet = workbook.createSheet("Сводка по предметам");
        Row header = sheet.createRow(0);
        createHeaderCell(header, 0, "Предмет", getTableHeaderStyle(workbook));
        createHeaderCell(header, 1, "Классов в сравнении", getTableHeaderStyle(workbook));
        createHeaderCell(header, 2, "Лист", getTableHeaderStyle(workbook));

        int rowNum = 1;
        for (Map.Entry<SubjectComparisonGroup, String> entry : groups.entrySet()) {
            SubjectComparisonGroup group = entry.getKey();
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(group.subject());
            row.createCell(1).setCellValue(group.classes().size());
            row.createCell(2).setCellValue(entry.getValue());
        }

        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createSubjectComparisonSheet(XSSFWorkbook workbook, SubjectComparisonGroup group, String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle headerStyle = getTableHeaderStyle(workbook);
        CellStyle centeredStyle = getStyle(workbook, StyleType.CENTERED);
        CellStyle percentStyle = getStyle(workbook, StyleType.PERCENT);

        Row title = sheet.createRow(0);
        Cell titleCell = title.createCell(0);
        titleCell.setCellValue("Сравнение по предмету: " + group.subject());
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));

        Row header = sheet.createRow(2);
        createHeaderCell(header, 0, "Класс", headerStyle);
        createHeaderCell(header, 1, "ЕГКР 1 (дата)", headerStyle);
        createHeaderCell(header, 2, "ЕГКР 1 %", headerStyle);
        createHeaderCell(header, 3, "ЕГКР 2 (дата)", headerStyle);
        createHeaderCell(header, 4, "ЕГКР 2 %", headerStyle);
        createHeaderCell(header, 5, "ЕГЭ (дата)", headerStyle);
        createHeaderCell(header, 6, "ЕГЭ %", headerStyle);

        int rowNum = 3;
        for (ClassComparison classComparison : group.classes()) {
            Row row = sheet.createRow(rowNum++);
            setStyledValue(row, 0, classComparison.className(), centeredStyle);
            setStyledValue(row, 1, formatDate(classComparison.egkr1().getTestDate()), centeredStyle);
            setStyledValue(row, 2, safePercent(classComparison.egkr1().getSuccessPercentage()), percentStyle);
            setStyledValue(row, 3, formatDate(classComparison.egkr2().getTestDate()), centeredStyle);
            setStyledValue(row, 4, safePercent(classComparison.egkr2().getSuccessPercentage()), percentStyle);
            setStyledValue(row, 5, classComparison.ege() != null ? formatDate(classComparison.ege().getTestDate()) : "", centeredStyle);
            setStyledValue(row, 6, classComparison.ege() != null
                    ? safePercent(classComparison.ege().getSuccessPercentage())
                    : 0.0, percentStyle);
        }

        int averageRowNum = rowNum + 1;
        Row avgHeader = sheet.createRow(averageRowNum);
        createHeaderCell(avgHeader, 0, "Средние значения", headerStyle);
        createHeaderCell(avgHeader, 1, "ЕГКР 1", headerStyle);
        createHeaderCell(avgHeader, 2, "ЕГКР 2", headerStyle);
        createHeaderCell(avgHeader, 3, "ЕГЭ", headerStyle);

        Row avgValues = sheet.createRow(averageRowNum + 1);
        setStyledValue(avgValues, 0, "Средний % выполнения", centeredStyle);
        setStyledValue(avgValues, 1, averageEgkr1(group.classes()) / 100.0, percentStyle);
        setStyledValue(avgValues, 2, averageEgkr2(group.classes()) / 100.0, percentStyle);
        setStyledValue(avgValues, 3, averageEge(group.classes()) / 100.0, percentStyle);

        createClassComparisonChart(sheet, 2, rowNum - 1);
        createAverageComparisonChart(sheet, averageRowNum + 1);

        for (int i = 0; i < 8; i++) {
            sheet.autoSizeColumn(i);
        }
        applyPrintLayout(sheet);
    }

    private void createClassComparisonChart(Sheet sheet, int startHeaderRow, int endDataRow) {
        if (endDataRow <= startHeaderRow) {
            return;
        }

        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 8, 2, 15, 18));
        chart.setTitleText("ЕГКР 1 и ЕГКР 2 по классам");
        chart.setTitleOverlay(false);
        chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Класс");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("% выполнения");

        XDDFDataSource<String> classes = XDDFDataSourcesFactory.fromStringCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new CellRangeAddress(startHeaderRow + 1, endDataRow, 0, 0)
        );

        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
        XDDFNumericalDataSource<Double> egkr1 = XDDFDataSourcesFactory.fromNumericCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new CellRangeAddress(startHeaderRow + 1, endDataRow, 2, 2)
        );
        XDDFNumericalDataSource<Double> egkr2 = XDDFDataSourcesFactory.fromNumericCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new CellRangeAddress(startHeaderRow + 1, endDataRow, 4, 4)
        );

        XDDFLineChartData.Series s1 = (XDDFLineChartData.Series) data.addSeries(classes, egkr1);
        s1.setTitle("ЕГКР 1", null);
        s1.setMarkerStyle(MarkerStyle.CIRCLE);
        applySeriesColor(s1, 0);

        XDDFLineChartData.Series s2 = (XDDFLineChartData.Series) data.addSeries(classes, egkr2);
        s2.setTitle("ЕГКР 2", null);
        s2.setMarkerStyle(MarkerStyle.CIRCLE);
        applySeriesColor(s2, 1);

        chart.plot(data);
    }

    private void createAverageComparisonChart(Sheet sheet, int averageValueRow) {
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 8, 19, 15, 32));
        chart.setTitleText("Средние ЕГКР 1 / ЕГКР 2 / ЕГЭ");
        chart.setTitleOverlay(false);
        chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("% выполнения");

        if (sheet.getRow(averageValueRow - 1) == null) {
            return;
        }

        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new CellRangeAddress(averageValueRow - 1, averageValueRow - 1, 1, 3)
        );
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new CellRangeAddress(averageValueRow, averageValueRow, 1, 3)
        );

        XDDFBarChartData data = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        data.setBarDirection(BarDirection.COL);
        XDDFBarChartData.Series series = (XDDFBarChartData.Series) data.addSeries(categories, values);
        series.setTitle("Средний %", null);
        chart.plot(data);
    }

    private double safePercent(Double value) {
        return value == null ? 0.0 : value / 100.0;
    }

    private double averageEgkr1(List<ClassComparison> classes) {
        return classes.stream()
                .map(ClassComparison::egkr1)
                .map(TestSummaryDto::getSuccessPercentage)
                .filter(Objects::nonNull)
                .mapToDouble(Double::doubleValue)
                .average()
                .orElse(0.0);
    }

    private double averageEgkr2(List<ClassComparison> classes) {
        return classes.stream()
                .map(ClassComparison::egkr2)
                .map(TestSummaryDto::getSuccessPercentage)
                .filter(Objects::nonNull)
                .mapToDouble(Double::doubleValue)
                .average()
                .orElse(0.0);
    }

    private double averageEge(List<ClassComparison> classes) {
        return classes.stream()
                .map(ClassComparison::ege)
                .filter(Objects::nonNull)
                .map(TestSummaryDto::getSuccessPercentage)
                .filter(Objects::nonNull)
                .mapToDouble(Double::doubleValue)
                .average()
                .orElse(0.0);
    }

    private List<ComparisonGroup> buildGroups(List<TestSummaryDto> allTests) {
        List<TestSummaryDto> egkrTests = allTests.stream()
                .filter(t -> t.getTestType() != null && t.getTestType().startsWith("ЕГКР"))
                .filter(t -> t.getSubject() != null && t.getClassName() != null)
                .collect(Collectors.toList());

        Map<String, List<TestSummaryDto>> bySubjectClass = egkrTests.stream()
                .collect(Collectors.groupingBy(t -> t.getSubject() + "||" + t.getClassName()));

        List<ComparisonGroup> result = new ArrayList<>();

        for (Map.Entry<String, List<TestSummaryDto>> entry : bySubjectClass.entrySet()) {
            List<TestSummaryDto> groupTests = entry.getValue().stream()
                    .sorted(Comparator.comparing(TestSummaryDto::getTestDate))
                    .collect(Collectors.toList());

            if (groupTests.size() < 2) {
                continue;
            }

            TestSummaryDto latest = groupTests.get(groupTests.size() - 1);
            String teacher = latest.getTeacher();

            List<TestSummaryDto> teacherEgkrTests = groupTests.stream()
                    .filter(t -> Objects.equals(teacher, t.getTeacher()))
                    .sorted(Comparator.comparing(TestSummaryDto::getTestDate))
                    .collect(Collectors.toList());

            if (teacherEgkrTests.size() < 2) {
                continue;
            }

            List<TestSummaryDto> selected = new ArrayList<>(
                    teacherEgkrTests.subList(teacherEgkrTests.size() - 2, teacherEgkrTests.size())
            );

            Optional<TestSummaryDto> egeTest = allTests.stream()
                    .filter(t -> Objects.equals(t.getSubject(), latest.getSubject()))
                    .filter(t -> Objects.equals(t.getClassName(), latest.getClassName()))
                    .filter(t -> Objects.equals(t.getTeacher(), teacher))
                    .filter(t -> t.getTestType() != null && t.getTestType().contains("ЕГЭ"))
                    .max(Comparator.comparing(TestSummaryDto::getTestDate));

            egeTest.ifPresent(selected::add);

            result.add(new ComparisonGroup(
                    latest.getSubject(),
                    latest.getClassName(),
                    teacher,
                    selected
            ));
        }

        return result;
    }

    private Map<ComparisonGroup, String> buildUniqueSheetNames(List<ComparisonGroup> groups) {
        Map<ComparisonGroup, String> names = new LinkedHashMap<>();
        Set<String> usedNames = new HashSet<>();

        for (ComparisonGroup group : groups) {
            String base = group.sheetName();
            String candidate = base;
            int suffix = 2;

            while (usedNames.contains(candidate)) {
                String suffixPart = "_" + suffix++;
                int maxBaseLength = Math.max(1, 31 - suffixPart.length());
                String truncatedBase = base.length() > maxBaseLength ? base.substring(0, maxBaseLength) : base;
                candidate = truncatedBase + suffixPart;
            }

            usedNames.add(candidate);
            names.put(group, candidate);
        }

        return names;
    }

    private void createSummarySheet(XSSFWorkbook workbook, Map<ComparisonGroup, String> groups) {
        Sheet sheet = workbook.createSheet("Сводка сравнения");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Предмет");
        header.createCell(1).setCellValue("Класс");
        header.createCell(2).setCellValue("Учитель");
        header.createCell(3).setCellValue("Работ в сравнении");
        header.createCell(4).setCellValue("Лист");

        int rowNum = 1;
        for (Map.Entry<ComparisonGroup, String> entry : groups.entrySet()) {
            ComparisonGroup group = entry.getKey();
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(group.subject());
            row.createCell(1).setCellValue(group.className());
            row.createCell(2).setCellValue(group.teacher());
            row.createCell(3).setCellValue(group.tests().size());
            row.createCell(4).setCellValue(entry.getValue());
        }

        for (int i = 0; i < 5; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createComparisonSheet(XSSFWorkbook workbook, ComparisonGroup group, String sheetName) {
        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle headerStyle = getTableHeaderStyle(workbook);
        CellStyle percentStyle = getStyle(workbook, StyleType.PERCENT);
        CellStyle decimalStyle = getStyle(workbook, StyleType.DECIMAL);
        CellStyle centeredStyle = getStyle(workbook, StyleType.CENTERED);

        Row title = sheet.createRow(0);
        Cell titleCell = title.createCell(0);
        titleCell.setCellValue(
                String.format("Сравнение: %s, %s (%s)", group.subject(), group.className(), group.teacher()));
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));

        Row header = sheet.createRow(2);
        createHeaderCell(header, 0, "Дата", headerStyle);
        createHeaderCell(header, 1, "Тип", headerStyle);
        createHeaderCell(header, 2, "Учитель", headerStyle);
        createHeaderCell(header, 3, "% выполнения", headerStyle);
        createHeaderCell(header, 4, "Средний балл", headerStyle);
        createHeaderCell(header, 5, "% присутствия", headerStyle);
        createHeaderCell(header, 6, "Участников", headerStyle);

        int rowNum = 3;
        for (TestSummaryDto test : group.tests()) {
            Row row = sheet.createRow(rowNum++);
            setStyledValue(row, 0, formatDate(test.getTestDate()), centeredStyle);
            setStyledValue(row, 1, nullSafe(test.getTestType()), centeredStyle);
            setStyledValue(row, 2, nullSafe(test.getTeacher()), centeredStyle);
            setStyledValue(row, 3, test.getSuccessPercentage() / 100.0, percentStyle);
            setStyledValue(row, 4, test.getAverageScore() != null ? test.getAverageScore() : 0.0, decimalStyle);
            setStyledValue(row, 5, test.getAttendancePercentage() / 100.0, percentStyle);
            setStyledValue(row, 6, test.getStudentsPresent() != null ? test.getStudentsPresent() : 0, centeredStyle);
        }

        int legendRow = rowNum + 1;
        createLegendRow(sheet, legendRow, group.tests(), workbook);
        int tableStart = legendRow + 2;
        Row taskHeader = sheet.createRow(tableStart);
        createHeaderCell(taskHeader, 0, "Задание", headerStyle);
        for (int i = 0; i < group.tests().size(); i++) {
            TestSummaryDto test = group.tests().get(i);
            createHeaderCell(taskHeader, i + 1,
                    String.format("%s (%s)", nullSafe(test.getTestType()), formatDate(test.getTestDate())),
                    headerStyle);
        }

        Map<Integer, List<Double>> taskData = buildTaskComparisonData(group.tests());
        int row = tableStart + 1;
        for (Map.Entry<Integer, List<Double>> entry : taskData.entrySet()) {
            Row r = sheet.createRow(row++);
            setStyledValue(r, 0, entry.getKey(), centeredStyle);
            for (int i = 0; i < entry.getValue().size(); i++) {
                setStyledValue(r, i + 1, entry.getValue().get(i) / 100.0, percentStyle);
            }
        }

        if (!taskData.isEmpty()) {
            createLineChart(sheet, tableStart, row - 1, group.tests().size());
        }

        for (int i = 0; i < 9; i++) {
            sheet.autoSizeColumn(i);
        }
        applyPrintLayout(sheet);
    }

    private Map<Integer, List<Double>> buildTaskComparisonData(List<TestSummaryDto> tests) {
        List<Map<Integer, TaskStatisticsDto>> statsByTest = tests.stream()
                .map(t -> analysisService.getTaskStatistics(t.getReportFileId()))
                .collect(Collectors.toList());

        Set<Integer> allTasks = new TreeSet<>();
        for (Map<Integer, TaskStatisticsDto> map : statsByTest) {
            allTasks.addAll(map.keySet());
        }

        Map<Integer, List<Double>> result = new LinkedHashMap<>();
        for (Integer task : allTasks) {
            List<Double> values = new ArrayList<>();
            for (Map<Integer, TaskStatisticsDto> map : statsByTest) {
                TaskStatisticsDto dto = map.get(task);
                values.add(dto != null ? dto.getCompletionPercentage() : 0.0);
            }
            result.put(task, values);
        }

        return result;
    }

    private void createLineChart(Sheet sheet, int startHeaderRow, int endDataRow, int seriesCount) {
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 0, endDataRow + 2, 12, endDataRow + 20));
        chart.setTitleText("Сравнение выполнения заданий");
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Задание");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("% выполнения");
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFDataSource<Double> taskNumbers = XDDFDataSourcesFactory.fromNumericCellRange(
                (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                new org.apache.poi.ss.util.CellRangeAddress(startHeaderRow + 1, endDataRow, 0, 0)
        );

        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
        for (int i = 1; i <= seriesCount; i++) {
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                    (org.apache.poi.xssf.usermodel.XSSFSheet) sheet,
                    new org.apache.poi.ss.util.CellRangeAddress(startHeaderRow + 1, endDataRow, i, i)
            );
            XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(taskNumbers, values);
            series.setTitle(sheet.getRow(startHeaderRow).getCell(i).getStringCellValue(), null);
            series.setSmooth(false);
            series.setMarkerStyle(MarkerStyle.CIRCLE);
            applySeriesColor(series, i - 1);
        }

        chart.plot(data);
    }

    private void applySeriesColor(XDDFLineChartData.Series series, int index) {
        byte[][] palette = {
                {(byte) 31, (byte) 119, (byte) 180},
                {(byte) 214, (byte) 39, (byte) 40},
                {(byte) 44, (byte) 160, (byte) 44}
        };
        byte[] rgb = palette[index % palette.length];
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(rgb));
        XDDFLineProperties lineProperties = new XDDFLineProperties();
        lineProperties.setFillProperties(fill);
        series.setLineProperties(lineProperties);
    }

    private void createHeaderCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void setStyledValue(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void setStyledValue(Row row, int col, double value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void setStyledValue(Row row, int col, int value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void createLegendRow(Sheet sheet, int rowNum, List<TestSummaryDto> tests, Workbook workbook) {
        Row legendTitle = sheet.createRow(rowNum);
        legendTitle.createCell(0).setCellValue("Легенда графика:");

        byte[][] palette = {
                {(byte) 31, (byte) 119, (byte) 180},
                {(byte) 214, (byte) 39, (byte) 40},
                {(byte) 44, (byte) 160, (byte) 44}
        };

        CellStyle textStyle = getStyle(workbook, StyleType.NORMAL);
        for (int i = 0; i < tests.size(); i++) {
            Row row = sheet.createRow(rowNum + i + 1);
            Cell colorCell = row.createCell(0);
            CellStyle colorStyle = workbook.createCellStyle();
            colorStyle.cloneStyleFrom(textStyle);
            colorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            byte[] rgb = palette[i % palette.length];
            colorStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF), null));
            colorCell.setCellStyle(colorStyle);

            row.createCell(1).setCellValue(
                    String.format("%s (%s)", tests.get(i).getTestType(), formatDate(tests.get(i).getTestDate())));
        }
    }

    private void applyPrintLayout(Sheet sheet) {
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setAutobreaks(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        sheet.createFreezePane(0, 3);
        sheet.setMargin(Sheet.LeftMargin, 0.3);
        sheet.setMargin(Sheet.RightMargin, 0.3);
        sheet.setMargin(Sheet.TopMargin, 0.5);
        sheet.setMargin(Sheet.BottomMargin, 0.5);
        sheet.getWorkbook().setPrintArea(
                sheet.getWorkbook().getSheetIndex(sheet),
                0, Math.max(8, sheet.getRow(2).getLastCellNum()),
                0, sheet.getLastRowNum() + 1
        );
    }

    private String formatDate(LocalDate date) {
        return date == null ? "" : date.toString();
    }

    private String nullSafe(String value) {
        return value == null ? "" : value;
    }

    private record ComparisonGroup(String subject, String className, String teacher, List<TestSummaryDto> tests) {
        String sheetName() {
            String raw = (subject + "_" + className).replaceAll("[\\\\/*\\[\\]:?]", "_");
            return raw.length() > 28 ? raw.substring(0, 28) : raw;
        }
    }

    private record ClassComparison(String className, TestSummaryDto egkr1, TestSummaryDto egkr2, TestSummaryDto ege) {
    }

    private record SubjectComparisonGroup(String subject, List<ClassComparison> classes) {
        String sheetName() {
            String raw = ("Предмет_" + subject).replaceAll("[\\\\/*\\[\\]:?]", "_");
            return raw.length() > 31 ? raw.substring(0, 31) : raw;
        }
    }
}
