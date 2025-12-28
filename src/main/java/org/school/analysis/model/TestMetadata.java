package org.school.analysis.model;

import lombok.Data;
import java.time.LocalDate;
import java.util.Map;

/**
 * Метаданные контрольной работы/тестирования
 */
@Data
public class TestMetadata {
    // Основная информация
    private String subject;                    // Предмет
    private String className;                  // Класс
    private String testName;                   // Название работы: "Входная контрольная"
    private LocalDate testDate;                // Дата проведения
    private String teacher;                    // Учитель
    private String school;                     // Школа: "ГБОУ 7"

    // Параметры теста
    private int taskCount;                     // Количество заданий
    private Map<Integer, Integer> maxScores;   // Макс. баллы по заданиям
    private int maxTotalScore;                 // Максимальный итоговый балл
    private int timeLimit;                     // Время на выполнение (минуты)

    // Критерии оценки
    private Map<String, Double> gradeRanges;   // {"5": 85.0, "4": 70.0, "3": 50.0}
    private String evaluationSystem;           // Система оценивания: "5-балльная"

    // Дополнительно
    private String curriculum;                 // Учебная программа
    private String topics;                     // Тематика
    private String testType;                   // Тип: "Входной", "Промежуточный", "Итоговый"
    private String comment;                    // Комментарий

    /**
     * Рассчитать максимальный итоговый балл
     */
    public int calculateMaxTotalScore() {
        this.maxTotalScore = maxScores.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
        return this.maxTotalScore;
    }

    /**
     * Получить критерий для оценки
     */
    public double getGradeThreshold(String grade) {
        return gradeRanges.getOrDefault(grade, 0.0);
    }

    /**
     * Проверка валидности метаданных
     */
    public boolean isValid() {
        return subject != null &&
                className != null &&
                taskCount > 0 &&
                !maxScores.isEmpty();
    }
}