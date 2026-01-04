// src/main/java/org/school/analysis/service/ExcelReportService.java
package org.school.analysis.service;

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

}