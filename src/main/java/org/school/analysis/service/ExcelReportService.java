// src/main/java/org/school/analysis/service/ExcelReportService.java
package org.school.analysis.service;

import org.school.analysis.model.dto.TestResultsDto;
import org.school.analysis.model.dto.TestSummaryDto;

import java.io.File;
import java.util.List;

/**
 * Сервис для генерации Excel-отчетов
 */
public interface ExcelReportService {

    /**
     * Генерирует сводный отчет по всем тестам в Excel
     * @param tests список тестов из getAllTestsSummary
     * @return путь к созданному файлу
     */
    File generateSummaryReport(List<TestSummaryDto> tests);

    /**
     * Генерирует детальный отчет для учителя по конкретному тесту
     * @param testResults результаты теста
     * @param includeChart включать ли диаграмму
     * @return путь к созданному файлу
     */
    File generateTeacherReport(TestResultsDto testResults, boolean includeChart);

    /**
     * Генерирует отчет с диаграммой по заданиям
     * @param testResults результаты теста
     * @return путь к созданному файлу
     */
    File generateTaskAnalysisReport(TestResultsDto testResults);

    /**
     * Генерирует все отчеты (сводный + детальный) по списку тестов
     * @param tests список тестов
     * @return список созданных файлов
     */
    List<File> generateAllReports(List<TestSummaryDto> tests);
}