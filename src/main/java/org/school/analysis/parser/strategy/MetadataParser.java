package org.school.analysis.parser.strategy;

import org.apache.poi.ss.usermodel.*;
import org.school.analysis.model.TestMetadata;
import org.school.analysis.util.ExcelParser;
import org.springframework.stereotype.Component;

import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

/**
 * Стратегия парсинга метаданных из Excel файла
 */
@Component
public class MetadataParser {

    /**
     * Парсинг метаданных из листа "Информация"
     */
    public TestMetadata parseMetadata(Sheet infoSheet) {
        TestMetadata metadata = new TestMetadata();

        if (infoSheet == null) {
            return metadata;
        }

        // Основная информация (строки 1-5)
        metadata.setTeacher(ExcelParser.getCellValueAsString(infoSheet, 0, 1, "Не указан"));
        metadata.setTestDate(parseDate(ExcelParser.getCellValueAsString(infoSheet, 1, 1)));
        metadata.setSubject(ExcelParser.getCellValueAsString(infoSheet, 2, 1, "Неизвестный предмет"));
        metadata.setClassName(ExcelParser.getCellValueAsString(infoSheet, 3, 1, "Неизвестный класс"));
        metadata.setTestType(ExcelParser.getCellValueAsString(infoSheet, 4, 1, "Неизвествный тип работы"));
        metadata.setSchool(ExcelParser.getCellValueAsString(infoSheet, 8, 1, "ГБОУ №7"));

        // Парсинг строки с максимальными баллами (если есть)
        String scoresText = ExcelParser.getCellValueAsString(infoSheet, 5, 1);
        if (scoresText != null) {
            metadata.setMaxScores(parseMaxScoresFromText(scoresText));
        }


        return metadata;
    }

    /**
     * Парсинг максимальных баллов из текста
     * Формат: "1=2, 2=2, 3=3, 4=1, 5=2"
     */
    private Map<Integer, Integer> parseMaxScoresFromText(String text) {
        Map<Integer, Integer> maxScores = new HashMap<>();

        if (text == null || text.trim().isEmpty()) {
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

    /**
     * Критерии оценок по умолчанию
     */
    private Map<String, Double> getDefaultGradeRanges() {
        Map<String, Double> ranges = new HashMap<>();
        ranges.put("5", 85.0);  // от 85%
        ranges.put("4", 70.0);  // от 70%
        ranges.put("3", 50.0);  // от 50%
        ranges.put("2", 0.0);   // менее 50%
        return ranges;
    }

    /**
     * Извлечение метаданных из имени файла
     */
    public TestMetadata parseFromFileName(String fileName) {
        TestMetadata metadata = new TestMetadata();

        // Пример: "Сбор_данных_10А_История_2025-01-15.xlsx"
        String nameWithoutExt = fileName.replace(".xlsx", "");
        String[] parts = nameWithoutExt.split("_");

        for (String part : parts) {
            if (part.matches("\\d+[А-Яа-я]?")) {
                metadata.setClassName(part); // Класс
            } else if (isDate(part)) {
                metadata.setTestDate(LocalDate.parse(part)); // Дата
            } else if (!part.equalsIgnoreCase("Сбор") &&
                    !part.equalsIgnoreCase("данных") &&
                    !part.equalsIgnoreCase("класс")) {
                metadata.setSubject(part); // Предмет
            }
        }

        return metadata;
    }

    private boolean isDate(String str) {
        try {
            LocalDate.parse(str);
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}