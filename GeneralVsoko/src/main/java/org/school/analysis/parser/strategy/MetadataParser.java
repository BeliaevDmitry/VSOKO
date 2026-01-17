package org.school.analysis.parser.strategy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.school.analysis.exception.ValidationException;
import org.school.analysis.model.TestMetadata;
import org.school.analysis.service.TeacherService;
import org.school.analysis.util.ExcelParser;
import org.school.analysis.util.ValidationHelper;
import org.springframework.stereotype.Component;

import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

/**
 * Стратегия парсинга метаданных из Excel файла
 */
@Slf4j
@Component
public class MetadataParser {

    private final TeacherService teacherService;

    public MetadataParser(TeacherService teacherService) {
        this.teacherService = teacherService;
    }

    public TestMetadata parseMetadata(Sheet infoSheet) {
        TestMetadata metadata = new TestMetadata();

        if (infoSheet == null) {
            throw new ValidationException("Лист 'Информация' не найден");
        }

        try {
            // 1. Собираем все данные
            String rawTeacherName = ExcelParser.getCellValueAsString(infoSheet, 0, 1, "Не указан");
            String testDateStr = ExcelParser.getCellValueAsString(infoSheet, 1, 1);
            String subject = ExcelParser.getCellValueAsString(infoSheet, 2, 1, "Неизвестный предмет");
            String className = ExcelParser.getCellValueAsString(infoSheet, 3, 1, "Неизвестный класс");
            String maxScoresText = ExcelParser.getCellValueAsString(infoSheet, 5, 1, "нет баллов");

            // 2. Проверяем обязательные поля
            ValidationHelper.ValidationResult validation = validateMandatoryFields(
                    rawTeacherName, testDateStr, subject, className);

            if (validation.hasErrors()) {
                throw new ValidationException(validation.getErrorsAsString());
            }

            // 3. Создаем метаданные
            metadata.setTeacher(getCorrectTeacherName(rawTeacherName));
            metadata.setTestDate(parseDate(testDateStr));
            metadata.setSubject(subject);
            metadata.setClassName(className);
            metadata.setTestType(ExcelParser.getCellValueAsString(infoSheet, 4, 1, "Неизвестный тип работы"));
            metadata.setMaxScores(parseMaxScoresFromText(maxScoresText));
            metadata.setComment(ExcelParser.getCellValueAsString(infoSheet, 6, 1, ""));
            metadata.setSchoolName(ExcelParser.getCellValueAsString(infoSheet, 7, 1, "ГБОУ №7"));
            metadata.setAcademicYear(ExcelParser.getCellValueAsString(infoSheet, 8, 1, "2025-2026"));

            return metadata;

        } catch (Exception e) {
            if (e instanceof ValidationException) {
                throw e;
            }
            throw new ValidationException("Ошибка парсинга метаданных: " + e.getMessage(), e);
        }
    }

    private ValidationHelper.ValidationResult validateMandatoryFields(
            String teacherName, String testDateStr, String subject, String className) {

        ValidationHelper.ValidationResult result = new ValidationHelper.ValidationResult();

        // Проверка учителя
        ValidationHelper.ValidationResult teacherValidation =
                ValidationHelper.validateTeacher(teacherName, teacherService);
        if (teacherValidation.hasErrors()) {
            result.addError("Учитель: " + teacherValidation.getErrorsAsString());
        }

        // Проверка даты
        if (testDateStr == null || testDateStr.trim().isEmpty()) {
            result.addError("Не указана дата теста");
        }

        // Проверка предмета
        if (!ValidationHelper.isValidSubject(subject)) {
            result.addError("Некорректный предмет: " + subject);
        }

        // Проверка класса
        if (!ValidationHelper.isValidClassName(className)) {
            result.addError("Некорректный класс: " + className);
        }

        return result;
    }

    /**
     * Получает корректное полное ФИО учителя из базы данных
     */
    private String getCorrectTeacherName(String rawTeacherName) {
        // Если учитель не указан, возвращаем как есть
        if ("Не указан".equals(rawTeacherName) ||
                rawTeacherName == null ||
                rawTeacherName.trim().isEmpty()) {
            return "Не указан";
        }

        // Пытаемся получить полное ФИО из базы
        Optional<String> fullNameOpt = teacherService.getFullTeacherName(rawTeacherName);

        if (fullNameOpt.isPresent()) {
            // Возвращаем полное ФИО из базы
            return fullNameOpt.get();
        } else {
            // Эта ситуация НЕ должна возникать, т.к. validateMandatoryFields уже проверила
            // Но на всякий случай возвращаем исходное имя
            log.warn("Учитель '{}' прошел валидацию, но не найден в базе для получения полного имени",
                    rawTeacherName);
            return rawTeacherName;
        }
    }

    /**
     * Парсинг максимальных баллов из текста
     * Формат: "1=2, 2=2, 3=3, 4=1, 5=2"
     */
    private Map<Integer, Integer> parseMaxScoresFromText(String text) {
        Map<Integer, Integer> maxScores = new HashMap<>();

        if (text == null || text.trim().isEmpty() || "нет баллов".equalsIgnoreCase(text.trim())) {
            return maxScores;
        }

        String[] pairs = text.split(",");
        for (String pair : pairs) {
            String[] keyValue = pair.trim().split("=");
            if (keyValue.length == 2) {
                try {
                    int taskNum = Integer.parseInt(keyValue[0].trim());
                    int maxScore = Integer.parseInt(keyValue[1].trim());
                    maxScores.put(taskNum, maxScore);
                } catch (NumberFormatException e) {
                    // Пропускаем некорректные пары
                }
            }
        }

        return maxScores;
    }

    /**
     * Парсинг даты из строки
     */
    private LocalDate parseDate(String dateString) {
        if (dateString == null || dateString.trim().isEmpty()) {
            return LocalDate.now();
        }

        try {
            // Попробуем разные форматы
            if (dateString.contains(".")) {
                String[] parts = dateString.split("\\.");
                if (parts.length == 3) {
                    return LocalDate.of(
                            Integer.parseInt(parts[2]),
                            Integer.parseInt(parts[1]),
                            Integer.parseInt(parts[0])
                    );
                }
            }
            return LocalDate.parse(dateString);
        } catch (Exception e) {
            return LocalDate.now();
        }
    }
}