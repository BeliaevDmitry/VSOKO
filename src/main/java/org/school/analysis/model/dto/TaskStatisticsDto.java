package org.school.analysis.model.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TaskStatisticsDto {
    private Integer taskNumber;
    private Integer maxScore;
    private Integer averageScore;
    private Integer completedCount; // Сколько учеников выполнили полностью
    private Integer partiallyCompletedCount; // Выполнили частично
    private Integer notCompletedCount; // Не выполнили
    private Double completionPercentage; // Процент выполнения
}