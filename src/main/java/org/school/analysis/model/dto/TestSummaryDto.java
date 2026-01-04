// src/main/java/org/school/analysis/model/dto/AnalysisDto.java
package org.school.analysis.model.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TestSummaryDto {
    private String subject;
    private String className;
    private LocalDate testDate;
    private String teacher;
    private String school;
    private Integer studentCount;
    private Integer taskCount;
    private Integer maxTotalScore;
    private String testType;
    private String fileName;
}



