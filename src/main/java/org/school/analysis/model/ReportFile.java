package org.school.analysis.model;

import lombok.Data;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Map;

@Data
public class ReportFile {
    private File file;                    // Файл
    private String subject;               // Предмет (из файла или имени)
    private String className;             // Класс (из файла или имени)
    private ProcessingStatus status;      // Статус обработки
    private LocalDateTime processedAt;    // Когда обработан
    private String errorMessage;          // Сообщение об ошибке (если есть)
    private int studentCount;             // Количество учеников в файле
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

    public String getFileName() {
        return file != null ? file.getName() : "unknown";
    }
}