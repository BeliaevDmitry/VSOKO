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
    private LocalDate testDate;                // Дата проведения
    private String teacher;                    // Учитель
    private String school = "ГБОУ №7";          // Школа: "ГБОУ 7"

    // Параметры теста
    private int taskCount;                     // Количество заданий
    private Map<Integer, Integer> maxScores;   // Макс. баллы по заданиям
    private int maxTotalScore;                 // Максимальный итоговый балл

    // Дополнительно
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
     * Проверка валидности метаданных
     */
    public boolean isValid() {
        return subject != null &&
                className != null &&
                taskCount > 0 &&
                !maxScores.isEmpty();
    }
}