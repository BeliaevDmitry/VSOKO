package org.school.analysis.model;

import lombok.Data;
import java.time.LocalDate;
import java.util.Map;

/**
 * Временный DTO для парсинга метаданных из Excel
 * После парсинга данные переносятся в ReportFile
 */
@Data
public class TestMetadata {
    // Основная информация
    private String subject;
    private String className;
    private LocalDate testDate;
    private String teacher;
    private String school = "ГБОУ №7";

    // Параметры теста (устанавливаются позже из данных учеников)
    private Map<Integer, Integer> maxScores;
    private int taskCount;
    private int maxTotalScore;

    // Дополнительно
    private String testType;
    private String comment;

    /**
     * Рассчитать максимальный итоговый балл
     */
    public int calculateMaxTotalScore() {
        if (maxScores == null || maxScores.isEmpty()) {
            this.maxTotalScore = 0;
            return 0;
        }
        this.maxTotalScore = maxScores.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
        return this.maxTotalScore;
    }
}