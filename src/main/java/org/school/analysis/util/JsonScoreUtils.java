package org.school.analysis.util;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.experimental.UtilityClass;

import java.util.HashMap;
import java.util.Map;

@UtilityClass
public class JsonScoreUtils {

    private static final ObjectMapper objectMapper = new ObjectMapper();

    /**
     * Преобразовать Map в JSON строку
     */
    public static String mapToJson(Map<Integer, Integer> scores) {
        if (scores == null || scores.isEmpty()) {
            return null;
        }
        try {
            return objectMapper.writeValueAsString(scores);
        } catch (Exception e) {
            throw new RuntimeException("Ошибка сериализации баллов в JSON: " + e.getMessage(), e);
        }
    }

    /**
     * Преобразовать JSON строку в Map
     */
    public static Map<Integer, Integer> jsonToMap(String json) {
        if (json == null || json.trim().isEmpty()) {
            return new HashMap<>();
        }
        try {
            return objectMapper.readValue(
                    json,
                    new TypeReference<Map<Integer, Integer>>() {}
            );
        } catch (Exception e) {
            throw new RuntimeException("Ошибка парсинга JSON баллов: " + e.getMessage(), e);
        }
    }

    /**
     * Рассчитать сумму баллов из JSON
     */
    public static Integer calculateTotalScore(String json) {
        Map<Integer, Integer> scores = jsonToMap(json);
        return scores.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
    }

    /**
     * Рассчитать сумму баллов из Map
     */
    public static Integer calculateTotalScore(Map<Integer, Integer> scores) {
        if (scores == null || scores.isEmpty()) {
            return 0;
        }
        return scores.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
    }
}