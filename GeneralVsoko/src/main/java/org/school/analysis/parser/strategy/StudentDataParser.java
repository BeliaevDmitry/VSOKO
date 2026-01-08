package org.school.analysis.parser.strategy;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.school.analysis.exception.ValidationException;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.ExcelParser;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Component;

import java.util.*;

@Component
public class StudentDataParser {

    private static final Logger log = LoggerFactory.getLogger(StudentDataParser.class);

    /**
     * Парсинг данных учеников с листа "Сбор информации"
     */
    public List<StudentResult> parseStudentData(Sheet dataSheet,
                                                Map<Integer, Integer> maxScores,
                                                String subject,
                                                String className) {
        List<StudentResult> results = new ArrayList<>();
        List<String> validationErrors = new ArrayList<>();

        if (maxScores.isEmpty()) {
            log.error("Нет данных о максимальных баллах. Невозможно парсить учеников.");
            return results;
        }

        log.debug("Парсинг учеников. Максимальные баллы: {}", maxScores);
        log.debug("Всего строк в листе: {}", dataSheet.getLastRowNum() + 1);

        // Данные учеников начинаются с 4-й строки (индекс 3)
        int firstStudentRow = 3;

        for (int rowIdx = firstStudentRow; rowIdx <= dataSheet.getLastRowNum(); rowIdx++) {
            Row row = dataSheet.getRow(rowIdx);
            if (row == null) {
                log.debug("Строка {} пустая", rowIdx);
                continue;
            }

            StudentResult result = parseStudentRow(row, maxScores, subject, className, rowIdx);

            if (result != null) {
                results.add(result);
            }
        }

        log.info("Найдено {} учеников", results.size());
        return results;
    }

    /**
     * Парсинг одной строки с данными ученика
     */
    private StudentResult parseStudentRow(Row row,
                                          Map<Integer, Integer> maxScores,
                                          String subject,
                                          String className,
                                          int rowIndex) {
        // ФИО (колонка B, индекс 1)
        String fio = ExcelParser.getCellValueAsString(row.getCell(1));
        if (fio == null || fio.trim().isEmpty() || fio.equals("[ПУСТО]")) {
            log.debug("Строка {}: ФИО пустое или '[ПУСТО]'", rowIndex + 1);
            return null;
        }

        fio = fio.trim();

        // Присутствие (колонка C, индекс 2)
        String presence = ExcelParser.getCellValueAsString(row.getCell(2));
        boolean wasPresent = "Был".equalsIgnoreCase(presence) || "Была".equalsIgnoreCase(presence);

        // Вариант (колонка D, индекс 3)
        String variant = ExcelParser.getCellValueAsString(row.getCell(3));

        // Парсинг баллов за задания (начиная с колонки E, индекс 4)
        Map<Integer, Integer> taskScores = parseTaskScores(row, maxScores.keySet(), rowIndex);

        // Создание результата
        StudentResult result = new StudentResult();
        result.setFio(fio);
        result.setPresence(presence);
        result.setVariant(variant);
        result.setSubject(subject);
        result.setClassName(className);
        result.setTaskScores(taskScores);

        // Вычисляем totalScore сразу при парсинге
        if (taskScores != null && !taskScores.isEmpty()) {
            result.setTotalScore(JsonScoreUtils.calculateTotalScore(taskScores));
        }

        return result;
    }

    /**
     * Парсинг баллов за задания
     */
    private Map<Integer, Integer> parseTaskScores(Row row, Set<Integer> taskNumbers, int rowIndex) {
        Map<Integer, Integer> scores = new HashMap<>();

        for (Integer taskNum : taskNumbers) {
            // Начинаем с колонки E (индекс 4)
            int columnIndex = 4 + (taskNum - 1);
            Cell cell = row.getCell(columnIndex);

            Integer score = ExcelParser.getCellValueAsInteger(cell);
            if (score == null) {
                score = 0; // Если ячейка пустая
            }

            scores.put(taskNum, score);
        }

        return scores;
    }
}