package org.school.analysis.service.impl.report;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
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
                createSummarySheet(workbook, groups);

                for (ComparisonGroup group : groups) {
                    createComparisonSheet(workbook, group);
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

    private void createSummarySheet(XSSFWorkbook workbook, List<ComparisonGroup> groups) {
        Sheet sheet = workbook.createSheet("Сводка сравнения");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Предмет");
        header.createCell(1).setCellValue("Класс");
        header.createCell(2).setCellValue("Учитель");
        header.createCell(3).setCellValue("Работ в сравнении");
        header.createCell(4).setCellValue("Лист");

        int rowNum = 1;
        for (ComparisonGroup group : groups) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(group.subject());
            row.createCell(1).setCellValue(group.className());
            row.createCell(2).setCellValue(group.teacher());
            row.createCell(3).setCellValue(group.tests().size());
            row.createCell(4).setCellValue(group.sheetName());
        }

        for (int i = 0; i < 5; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createComparisonSheet(XSSFWorkbook workbook, ComparisonGroup group) {
        String sheetName = group.sheetName();
        Sheet sheet = workbook.createSheet(sheetName);

        Row title = sheet.createRow(0);
        title.createCell(0).setCellValue(
                String.format("Сравнение: %s, %s (%s)", group.subject(), group.className(), group.teacher()));

        Row header = sheet.createRow(2);
        header.createCell(0).setCellValue("Дата");
        header.createCell(1).setCellValue("Тип");
        header.createCell(2).setCellValue("Учитель");
        header.createCell(3).setCellValue("% выполнения");
        header.createCell(4).setCellValue("Средний балл");
        header.createCell(5).setCellValue("% присутствия");
        header.createCell(6).setCellValue("Участников");

        int rowNum = 3;
        for (TestSummaryDto test : group.tests()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(formatDate(test.getTestDate()));
            row.createCell(1).setCellValue(nullSafe(test.getTestType()));
            row.createCell(2).setCellValue(nullSafe(test.getTeacher()));
            row.createCell(3).setCellValue(test.getSuccessPercentage());
            row.createCell(4).setCellValue(test.getAverageScore() != null ? test.getAverageScore() : 0.0);
            row.createCell(5).setCellValue(test.getAttendancePercentage());
            row.createCell(6).setCellValue(test.getStudentsPresent() != null ? test.getStudentsPresent() : 0);
        }

        int tableStart = rowNum + 2;
        Row taskHeader = sheet.createRow(tableStart);
        taskHeader.createCell(0).setCellValue("Задание");
        for (int i = 0; i < group.tests().size(); i++) {
            TestSummaryDto test = group.tests().get(i);
            taskHeader.createCell(i + 1).setCellValue(
                    String.format("%s (%s)", nullSafe(test.getTestType()), formatDate(test.getTestDate())));
        }

        Map<Integer, List<Double>> taskData = buildTaskComparisonData(group.tests());
        int row = tableStart + 1;
        for (Map.Entry<Integer, List<Double>> entry : taskData.entrySet()) {
            Row r = sheet.createRow(row++);
            r.createCell(0).setCellValue(entry.getKey());
            for (int i = 0; i < entry.getValue().size(); i++) {
                r.createCell(i + 1).setCellValue(entry.getValue().get(i));
            }
        }

        if (!taskData.isEmpty()) {
            createLineChart(sheet, tableStart, row - 1, group.tests().size());
        }

        for (int i = 0; i < 8; i++) {
            sheet.autoSizeColumn(i);
        }
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
        }

        chart.plot(data);
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
}

