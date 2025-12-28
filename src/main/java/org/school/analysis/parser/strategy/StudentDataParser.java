package org.school.analysis.parser.strategy;

import org.apache.poi.ss.usermodel.*;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.ExcelParser;

import java.util.*;

/**
 * Стратегия парсинга данных учеников
 */
public class StudentDataParser {

    /**
     * Парсинг данных учеников с листа "Сбор информации"
     */
    public List<StudentResult> parseStudentData(Sheet dataSheet,
                                                Map<Integer, Integer> maxScores,
                                                String subject,
                                                String className) {
        List<StudentResult> results = new ArrayList<>();

        if (dataSheet == null || maxScores.isEmpty()) {
            return results;
        }

        int firstStudentRow = 3; // 4-я строка (после заголовков)
        int maxStudents = 34;

        for (int rowIdx = firstStudentRow;
             rowIdx < firstStudentRow + maxStudents;
             rowIdx++) {

            Row row = dataSheet.getRow(rowIdx);
            if (row == null) break;

            StudentResult result = parseStudentRow(row, maxScores, subject, className);
            if (result != null && result.wasPresent()) {
                results.add(result);
            }
        }

        return results;
    }

    /**
     * Парсинг одной строки с данными ученика
     */
    private StudentResult parseStudentRow(Row row,
                                          Map<Integer, Integer> maxScores,
                                          String subject,
                                          String className) {
        // ФИО (колонка B, индекс 1)
        String fio = ExcelParser.getCellValueAsString(row.getCell(1));
        if (fio == null || fio.trim().isEmpty()) {
            return null; // Пустая строка
        }

        // Присутствие (колонка C, индекс 2)
        String presence = ExcelParser.getCellValueAsString(row.getCell(2));
        if ("Не был".equalsIgnoreCase(presence)) {
            // Создаем запись для отсутствующего
            StudentResult absent = new StudentResult();
            absent.setFio(fio.trim());
            absent.setPresence(presence);
            absent.setSubject(subject);
            absent.setClassName(className);
            absent.setTaskScores(new HashMap<>());
            absent.setMaxScores(new HashMap<>(maxScores));
            return absent;
        }

        // Вариант (колонка D, индекс 3)
        String variant = ExcelParser.getCellValueAsString(row.getCell(3));

        // Парсинг баллов за задания (начиная с колонки E, индекс 4)
        Map<Integer, Integer> taskScores = parseTaskScores(row, maxScores.keySet());

        // Создание результата
        StudentResult result = new StudentResult();
        result.setFio(fio.trim());
        result.setPresence(presence);
        result.setVariant(variant);
        result.setSubject(subject);
        result.setClassName(className);
        result.setTaskScores(taskScores);
        result.setMaxScores(new HashMap<>(maxScores));
        result.calculateAll();

        return result;
    }

    /**
     * Парсинг баллов за задания
     */
    private Map<Integer, Integer> parseTaskScores(Row row, Set<Integer> taskNumbers) {
        Map<Integer, Integer> scores = new HashMap<>();

        for (Integer taskNum : taskNumbers) {
            int columnIndex = 3 + taskNum; // E=4, F=5, G=6, ...
            Cell cell = row.getCell(columnIndex);

            Integer score = ExcelParser.getCellValueAsInteger(cell);
            if (score == null) {
                score = 0; // Если ячейка пустая
            }

            scores.put(taskNum, score);
        }

        return scores;
    }

    /**
     * Чтение максимальных баллов из 3-й строки
     */
    public Map<Integer, Integer> parseMaxScores(Sheet dataSheet) {
        Map<Integer, Integer> maxScores = new HashMap<>();

        Row maxScoresRow = dataSheet.getRow(2); // 3-я строка
        if (maxScoresRow == null) {
            return maxScores;
        }

        int taskNumber = 1;
        for (int col = 4; col < 100; col++) { // Начиная с колонки E
            Cell cell = maxScoresRow.getCell(col);
            if (cell == null) break;

            Integer maxScore = ExcelParser.getCellValueAsInteger(cell);
            if (maxScore != null) {
                maxScores.put(taskNumber, maxScore);
                taskNumber++;
            } else {
                break;
            }
        }

        return maxScores;
    }

    /**
     * Получение списка номеров заданий из заголовка
     */
    public Set<Integer> getTaskNumbers(Sheet dataSheet) {
        Set<Integer> taskNumbers = new TreeSet<>();

        Row headerRow = dataSheet.getRow(1); // 2-я строка с номерами заданий
        if (headerRow == null) {
            return taskNumbers;
        }

        for (int col = 4; col < 100; col++) {
            Cell cell = headerRow.getCell(col);
            if (cell == null) break;

            String value = ExcelParser.getCellValueAsString(cell);
            if (value != null && value.matches("\\d+")) {
                taskNumbers.add(Integer.parseInt(value));
            } else {
                break;
            }
        }

        return taskNumbers;
    }
}