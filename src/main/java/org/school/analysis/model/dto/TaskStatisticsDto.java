package org.school.analysis.model.dto;

import lombok.Builder;
import lombok.Data;

import java.util.HashMap;
import java.util.Map;

@Data
@Builder
public class TaskStatisticsDto {
    private Integer taskNumber;
    private Integer maxScore;

    @Builder.Default
    private Map<Integer, Integer> scoreDistribution = new HashMap<>();

    public void incrementScoreCount(Integer score) {
        scoreDistribution.merge(score, 1, Integer::sum);
    }

    // Геттеры для удобства
    public int getFullyCompletedCount() {
        return scoreDistribution.getOrDefault(maxScore, 0);
    }

    public int getPartiallyCompletedCount() {
        return scoreDistribution.entrySet().stream()
                .filter(e -> e.getKey() > 0 && e.getKey() < maxScore)
                .mapToInt(Map.Entry::getValue)
                .sum();
    }

    public int getNotCompletedCount() {
        return scoreDistribution.getOrDefault(0, 0);
    }

    public int getTotalStudents() {
        return scoreDistribution.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
    }

    // Возвращает double (не может быть null)
    public double getCompletionPercentage() {
        int totalStudents = getTotalStudents();
        if (totalStudents == 0) return 0.0;

        double totalScoreSum = scoreDistribution.entrySet().stream()
                .mapToDouble(e -> e.getKey() * e.getValue())
                .sum();

        double maxPossibleSum = maxScore * totalStudents;
        if (maxPossibleSum == 0) return 0.0;

        return (totalScoreSum / maxPossibleSum) * 100.0;
    }
}