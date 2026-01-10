package org.school.analysis.parser.strategy;

import org.apache.poi.ss.usermodel.Sheet;
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

    public TestMetadata parseMetadata(Sheet infoSheet) {
        TestMetadata metadata = new TestMetadata();

        if (infoSheet == null) {
            return metadata;
        }

        // Основная информация
        metadata.setTeacher(ExcelParser.getCellValueAsString(infoSheet, 0, 1, "Не указан"));
        metadata.setTestDate(parseDate(ExcelParser.getCellValueAsString(infoSheet, 1, 1)));
        metadata.setSubject(ExcelParser.getCellValueAsString(infoSheet, 2, 1, "Неизвестный предмет"));
        metadata.setClassName(ExcelParser.getCellValueAsString(infoSheet, 3, 1, "Неизвестный класс"));
        metadata.setTestType(ExcelParser.getCellValueAsString(infoSheet, 4, 1, "Неизвестный тип работы"));
        metadata.setMaxScores(parseMaxScoresFromText(ExcelParser.getCellValueAsString(infoSheet, 5, 1, "нет баллов")));
        metadata.setComment(ExcelParser.getCellValueAsString(infoSheet, 6, 1, ""));
        metadata.setSchoolName(ExcelParser.getCellValueAsString(infoSheet, 7, 1, "ГБОУ №7"));
        metadata.setAcademicYear(ExcelParser.getCellValueAsString(infoSheet, 8, 1, "2025-2026"));
        return metadata;
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