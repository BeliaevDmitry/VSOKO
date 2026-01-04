package org.school.analysis.model.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TeacherReportDto {
    private String subject;
    private String className;
    private LocalDate testDate;
    private String teacher;
    private String testType;
    private List<StudentTestResultDto> studentResults;
    private Map<Integer, TaskStatisticsDto> taskStatistics; // Статистика по заданиям
}