package org.school.analysis.model.dto;

import lombok.Builder;
import lombok.Data;

import java.util.Map;

@Data
@Builder
public class StudentDetailedResultDto {
    private String fio;
    private String presence;
    private String variant;
    private Integer totalScore;
    private Double percentageScore;
    private Map<Integer, Integer> taskScores; // Баллы по заданиям
    private String ACADEMIC_YEAR;
    private String school;
}