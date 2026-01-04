package org.school.analysis.model.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TestResultsDto {
    private TestSummaryDto testSummary;
    private List<StudentTestResultDto> studentResults;
}