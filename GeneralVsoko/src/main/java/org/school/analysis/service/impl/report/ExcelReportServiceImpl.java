package org.school.analysis.service.impl.report;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.model.dto.*;
import org.school.analysis.service.ExcelReportService;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.List;
import java.util.Map;

@Service
@RequiredArgsConstructor
@Slf4j
public class ExcelReportServiceImpl implements ExcelReportService {

    private final SummaryReportGenerator summaryReportGenerator;
    private final DetailReportGenerator detailReportGenerator;
    private final TeacherReportGenerator teacherReportGenerator;

    @Override
    public File generateSummaryReport(List<TestSummaryDto> tests, String schoolName) {
        log.info("Генерация сводного отчета для {} тестов", tests.size());
        return summaryReportGenerator.generateSummaryReport(tests, schoolName);
    }

    @Override
    public File generateTestDetailReport(
            TestSummaryDto testSummary,
            List<StudentDetailedResultDto> studentResults,
            Map<Integer, TaskStatisticsDto> taskStatistics,
            String schoolName) {

        log.info("Генерация детального отчета для теста: {} - {}",
                testSummary.getSubject(), testSummary.getClassName());

        return detailReportGenerator.generateDetailReportFile(testSummary, studentResults, taskStatistics);
    }

    @Override
    public File generateTeacherReportWithDetails(
            String teacherName,
            List<TestSummaryDto> teacherTests,
            List<TeacherTestDetailDto> teacherTestDetails,
            String schoolName) {

        log.info("Генерация детального отчета для учителя: {} ({} тестов)",
                teacherName, teacherTests.size());

        return teacherReportGenerator.generateTeacherReportWithDetails(teacherName, teacherTests,
                teacherTestDetails, schoolName);
    }
}