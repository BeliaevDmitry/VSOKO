package org.school.analysis.model;

import lombok.Data;
import java.time.LocalDateTime;
import java.util.List;

/**
 * Результат парсинга файла с оценками.
 * Содержит либо данные учеников (если парсинг успешен),
 * либо информацию об ошибке (если парсинг не удался).
 */
@Data
public class ParseResult {
    /**
     * Файл, который парсили. Содержит путь, имя, предмет, класс.
     * Всегда заполнен.
     */
    private ReportFile reportFile;

    /**
     * Результаты учеников. Заполняется ТОЛЬКО при успешном парсинге.
     * Каждый StudentResult содержит баллы за задания, ФИО, вариант.
     * Если парсинг не удался - список будет пустым.
     */
    private List<StudentResult> studentResults;

    /**
     * Успешен ли парсинг файла:
     * true - файл распарсен, studentResults содержит данные
     * false - ошибка парсинга, смотри errorMessage
     */
    private boolean success;

    /**
     * Сообщение об ошибке. Заполняется ТОЛЬКО если success = false.
     * Примеры: "Неверный формат файла", "Файл поврежден", "Не найден лист с данными"
     */
    private String errorMessage;

    /**
     * Сколько учеников распарсено из файла.
     * При успехе: равно количеству элементов в studentResults
     * При ошибке: 0
     */
    private int parsedStudents;

    /**
     * Время когда файл был обработан парсером.
     * Автоматически устанавливается при создании ParseResult.
     */
    private LocalDateTime parsedAt;

    private List<String> warnings; // Добавляем для предупреждений

    // Конструкторы удалены - используйте статические фабричные методы ниже

    /**
     * Создает успешный результат парсинга.
     * @param reportFile файл который парсили
     * @param studentResults список результатов учеников (не null)
     * @return ParseResult с success = true
     */
    public static ParseResult success(ReportFile reportFile,
                                      List<StudentResult> studentResults) {
        if (studentResults == null) {
            throw new IllegalArgumentException("studentResults не может быть null");
        }

        ParseResult result = new ParseResult();
        result.setReportFile(reportFile);
        result.setStudentResults(studentResults);
        result.setSuccess(true);
        result.setParsedStudents(studentResults.size());
        result.setParsedAt(LocalDateTime.now());
        return result;
    }

    /**
     * Создает результат с ошибкой парсинга.
     * @param reportFile файл который не удалось распарсить
     * @param errorMessage описание ошибки (не null и не пустое)
     * @return ParseResult с success = false
     */
    public static ParseResult error(ReportFile reportFile, String errorMessage) {
        if (errorMessage == null || errorMessage.trim().isEmpty()) {
            throw new IllegalArgumentException("errorMessage не может быть пустым");
        }

        ParseResult result = new ParseResult();
        result.setReportFile(reportFile);
        result.setSuccess(false);
        result.setErrorMessage(errorMessage.trim());
        result.setParsedStudents(0);
        result.setStudentResults(List.of()); // пустой список вместо null
        result.setParsedAt(LocalDateTime.now());
        return result;
    }
}