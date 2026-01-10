package org.school.analysis.service;

import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.model.dto.TestSummaryDto;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Сервис для генерации Excel-отчетов
 */
public interface ExcelReportService {

    /**
     * Генерирует сводный отчет по всем тестам в Excel
     */
    File generateSummaryReport(List<TestSummaryDto> tests, String school);

    /**
     * Генерирует детальный отчет по тесту
     */
    File generateTestDetailReport(
            TestSummaryDto testSummary,
            List<StudentDetailedResultDto> studentResults,
            Map<Integer, TaskStatisticsDto> taskStatistics,
            String school);


    /**
     * Генерация отчета для учителя с детальными данными по тестам
     */
    File generateTeacherReportWithDetails(String teacherName,
                                          List<TestSummaryDto> teacherTests,
                                          List<TeacherTestDetailDto> teacherTestDetails,
                                          String school);
}