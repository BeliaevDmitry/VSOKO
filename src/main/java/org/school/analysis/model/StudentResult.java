package org.school.analysis.model;

import lombok.Data;

import java.time.LocalDate;
import java.util.Map;

/**
 * Полная модель результатов ученика с вычисляемыми полями
 */
@Data
public class StudentResult {
    // Основные данные
    private String subject;                    // Предмет: "История", "Математика"
    private String className;                  // Класс: "10А", "6Б", "11"
    private String fio;                        // ФИО: "Иванов Иван Иванович"
    private String presence;                   // Присутствие: "Был"/"Не был"
    private String variant;                    // Вариант: "Вариант 1", "Вариант 2"
    private String testType;                    // Тип: "Входной", "Промежуточный", "Итоговый"
    private LocalDate testDate;                // Дата проведения

    // Результаты
    private Map<Integer, Integer> taskScores;  // Баллы за задания: {1=2, 2=1, 3=0}

    
    /**
     * Получить балл за конкретное задание
     */
    public Integer getScoreForTask(int taskNumber) {
        return taskScores.get(taskNumber);
    }

    /**
     * Проверка, присутствовал ли ученик
     */
    public boolean wasPresent() {
        return "Был".equalsIgnoreCase(presence);
    }
}