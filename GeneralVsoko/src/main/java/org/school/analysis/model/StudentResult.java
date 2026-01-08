package org.school.analysis.model;

import lombok.Data;
import org.school.analysis.util.JsonScoreUtils;

import java.time.LocalDate;
import java.util.Map;

@Data
public class StudentResult {
    private String subject;
    private String className;
    private String fio;
    private String presence;
    private String variant;
    private String testType;
    private LocalDate testDate;
    private Integer totalScore;
    private Double percentageScore;
    private String school = "ГБОУ №7";
    private String ACADEMIC_YEAR = "2025-2026";

    // Map в памяти для удобной работы
    private Map<Integer, Integer> taskScores;

    // Для удобства - геттер JSON
    public String getTaskScoresJson() {
        return JsonScoreUtils.mapToJson(taskScores);
    }

    // Для удобства - сеттер из JSON
    public void setTaskScoresJson(String json) {
        this.taskScores = JsonScoreUtils.jsonToMap(json);
    }

    // Вычисляемые методы
    public boolean wasPresent() {
        return "Был".equalsIgnoreCase(presence) || "был".equalsIgnoreCase(presence);
    }

    public Integer getTotalScore() {
        if (totalScore != null) {
            return totalScore;
        }
        return JsonScoreUtils.calculateTotalScore(taskScores);
    }
}