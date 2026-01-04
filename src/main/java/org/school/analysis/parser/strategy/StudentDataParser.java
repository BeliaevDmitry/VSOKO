package org.school.analysis.parser.strategy;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.school.analysis.exception.ValidationException;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.ExcelParser;
import org.school.analysis.util.ValidationHelper;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Стратегия парсинга данных учеников
 */
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

            StudentResult result = parseStudentRow(row, maxScores, subject, className,
                    rowIdx);

            if (result != null) {
                // ВАЛИДАЦИЯ и проверка результата
                ValidationHelper.ValidationResult validation =
                        ValidationHelper.validateStudentResult(result, maxScores);

                if (!validation.isValid()) {
                    String errorMsg = String.format("Строка %d, ученик '%s': %s",
                            rowIdx + 1, result.getFio(),
                            String.join("; ", validation.getErrors()));
                    validationErrors.add(errorMsg);
                    log.warn(errorMsg);
                    continue;
                }

                // Проверяем предупреждения
                if (!validation.getWarnings().isEmpty()) {
                    for (String warning : validation.getWarnings()) {
                        log.warn("Предупреждение (строка {}): {}", rowIdx + 1, warning);
                    }
                }

                if (result.wasPresent()) {
                    // Сохраняем с баллами
                    results.add(result);
                    log.debug("Добавлен присутствовавший ученик: {} (баллы: {})",
                            result.getFio(), result.getTaskScores());
                } else {
                    // Сохраняем отсутствующего, но очищаем баллы
                    result.setTaskScores(new HashMap<>()); // или null
                    results.add(result);
                    log.debug("Добавлен отсутствовавший ученик: {}", result.getFio());
                }
            }
        }

        // Если есть критические ошибки, можно бросить исключение
        if (!validationErrors.isEmpty()) {
            throw new ValidationException("Ошибки валидации данных учеников:\n" +
                    String.join("\n", validationErrors));
        }

        log.info("Найдено {} учеников (присутствовавших)", results.size());
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
        log.debug("Парсинг ученика: {}", fio);

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
        return result;
    }

    /**
     * Парсинг баллов за задания
     */
    private Map<Integer, Integer> parseTaskScores(Row row, Set<Integer> taskNumbers, int rowIndex) {
        Map<Integer, Integer> scores = new HashMap<>();

        log.debug("Строка {}: парсинг баллов для заданий {}", rowIndex + 1, taskNumbers);

        for (Integer taskNum : taskNumbers) {
            // Начинаем с колонки E (индекс 4)
            // Задание 1 -> колонка E (индекс 4)
            // Задание 2 -> колонка F (индекс 5)
            // и т.д.
            int columnIndex = 4 + (taskNum - 1);
            Cell cell = row.getCell(columnIndex);

            Integer score = ExcelParser.getCellValueAsInteger(cell);
            if (score == null) {
                score = 0; // Если ячейка пустая
            }

            scores.put(taskNum, score);
            log.debug("  Задание {} (колонка {}): балл = {}", taskNum, columnIndex, score);
        }

        return scores;
    }

    /**
     * Логирование содержимого строки
     */
    private void logRowData(Row row) {
        if (row == null) {
            log.debug("  Строка пустая");
            return;
        }

        StringBuilder sb = new StringBuilder();
        sb.append("  ");

        int cellCount = 0;
        for (int i = 0; i <= row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                String value = ExcelParser.getCellValueAsString(cell);
                if (value != null && !value.isEmpty()) {
                    sb.append("[").append(i).append("]=").append(value).append(" ");
                    cellCount++;
                }
            }
        }

        if (cellCount == 0) {
            log.debug("  Нет данных в строке");
        } else {
            log.debug(sb.toString());
        }
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

        for (int col = 4; col <= headerRow.getLastCellNum(); col++) {
            Cell cell = headerRow.getCell(col);
            if (cell == null) break;

            String value = ExcelParser.getCellValueAsString(cell);
            if (value != null && value.matches("\\d+")) {
                taskNumbers.add(Integer.parseInt(value));
            } else {
                break;
            }
        }

        log.debug("Найдены номера заданий: {}", taskNumbers);
        return taskNumbers;
    }
}