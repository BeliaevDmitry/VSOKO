package org.school.analysis.model.dto;

import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * DTO для детальных данных теста в отчете учителя
 */
@Data
@Builder
public class TeacherTestDetailDto {
    private TestSummaryDto testSummary;
    private List<StudentDetailedResultDto> studentResults;
    private Map<Integer, TaskStatisticsDto> taskStatistics;

    // Вспомогательные методы для доступа к данным теста
    public String getReportFileId() {
        return testSummary != null ? testSummary.getReportFileId() : null;
    }

    public String getSubject() {
        return testSummary != null ? testSummary.getSubject() : null;
    }

    public String getClassName() {
        return testSummary != null ? testSummary.getClassName() : null;
    }

}