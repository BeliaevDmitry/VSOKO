package org.school.analysis.model;

import lombok.Data;
import org.school.analysis.util.JsonScoreUtils;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Map;

@Data
public class ReportFile {
    private File file;
    private String subject;
    private String className;
    private ProcessingStatus status;
    private LocalDateTime processedAt;
    private String errorMessage;
    private int studentCount;
    private LocalDate testDate;
    private String teacher;
    private String schoolName = "ГБОУ №7";
    private String academicYear = "2025-2026";

    // Параметры теста
    private int taskCount;
    private Map<Integer, Integer> maxScores;  // Теперь Map в памяти


    // Дополнительно
    private String testType;
    private String comment;

    public String getFileName() {
        return file != null ? file.getName() : "unknown";
    }

    /**
     * Получить максимальный балл как JSON для БД
     */
    public String getMaxScoresJson() {
        return JsonScoreUtils.mapToJson(maxScores);
    }

    /**
     * Установить максимальные баллы из JSON
     */
    public void setMaxScoresJson(String json) {
        this.maxScores = JsonScoreUtils.jsonToMap(json);
    }

    /**
     * Рассчитать максимальный итоговый балл
     */
    public int getMaxTotalScore() {
        return JsonScoreUtils.calculateTotalScore(maxScores);
    }

}