package org.school.analysis.model.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class StudentTestResultDto {
    private String fio;
    private String presence;
    private String variant;
    private Integer totalScore;
    private Double percentageScore;
    private Integer positionInClass; // место в классе (опционально)
}