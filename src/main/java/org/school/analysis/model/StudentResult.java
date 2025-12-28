package org.school.analysis.model;

import lombok.Data;
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

    // Результаты
    private Map<Integer, Integer> taskScores;  // Баллы за задания: {1=2, 2=1, 3=0}
    private Map<Integer, Integer> maxScores;   // Макс. баллы: {1=2, 2=2, 3=3}

    // Вычисляемые поля
    private int totalScore;                    // Сумма баллов: 3
    private int maxPossibleScore;              // Макс. возможный: 7
    private double percentage;                 // Процент: 42.86%
    private String grade;                      // Оценка: "3" (2-5)
    private int completedTasks;                // Выполнено заданий: 2 из 3

    /**
     * Рассчитать все вычисляемые поля
     */
    public void calculateAll() {
        this.totalScore = calculateTotalScore();
        this.maxPossibleScore = calculateMaxPossibleScore();
        this.percentage = calculatePercentage();
        this.grade = calculateGrade();
        this.completedTasks = calculateCompletedTasks();
    }

    private int calculateTotalScore() {
        return taskScores.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
    }

    private int calculateMaxPossibleScore() {
        return maxScores.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
    }

    private double calculatePercentage() {
        return maxPossibleScore > 0 ?
                (double) totalScore / maxPossibleScore * 100 : 0;
    }

    private String calculateGrade() {
        double percent = percentage;
        if (percent >= 85) return "5";
        if (percent >= 70) return "4";
        if (percent >= 50) return "3";
        return "2";
    }

    private int calculateCompletedTasks() {
        return (int) taskScores.entrySet().stream()
                .filter(entry -> {
                    Integer max = maxScores.get(entry.getKey());
                    return max != null && entry.getValue() > 0;
                })
                .count();
    }

    /**
     * Получить балл за конкретное задание
     */
    public Integer getScoreForTask(int taskNumber) {
        return taskScores.get(taskNumber);
    }

    /**
     * Получить максимальный балл за задание
     */
    public Integer getMaxScoreForTask(int taskNumber) {
        return maxScores.get(taskNumber);
    }

    /**
     * Проверка, присутствовал ли ученик
     */
    public boolean wasPresent() {
        return "Был".equalsIgnoreCase(presence);
    }
}