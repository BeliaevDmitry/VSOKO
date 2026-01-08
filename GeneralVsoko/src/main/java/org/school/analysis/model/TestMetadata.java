package org.school.analysis.model;

import lombok.Data;
import org.school.analysis.util.JsonScoreUtils;

import java.time.LocalDate;
import java.util.Map;

@Data
public class TestMetadata {
    private String subject;
    private String className;
    private LocalDate testDate;
    private String teacher;
    private String school = "ГБОУ №7";
    private Map<Integer, Integer> maxScores; // Map в памяти
    private String testType;
    private String comment;
    private String ACADEMIC_YEAR = "2025-2026";

    /**
     * Вычислить максимальный общий балл
     */
    public int getMaxTotalScore() {
        return JsonScoreUtils.calculateTotalScore(maxScores);
    }

    /**
     * Получить количество заданий
     */
    public int getTaskCount() {
        return maxScores != null ? maxScores.size() : 0;
    }
}