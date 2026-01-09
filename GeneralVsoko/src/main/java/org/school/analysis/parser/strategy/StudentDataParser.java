package org.school.analysis.parser.strategy;

import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.ExcelParser;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Component;

import java.util.*;

@Component
public class StudentDataParser {

    private static final Logger log = LoggerFactory.getLogger(StudentDataParser.class);

    // Константы для индексов колонок (на основе структуры файла)
    private static final int COL_FIO = 1;           // Колонка B - ФИО
    private static final int COL_PRESENCE = 2;      // Колонка C - Присутствие
    private static final int COL_VARIANT = 3;       // Колонка D - Вариант
    private static final int HEADER_ROW_INDEX = 1;  // Строка с номерами заданий (вторая строка)
    private static final int FIRST_STUDENT_ROW = 3; // Первая строка с данными студента
    private static final int MAX_NUMBER_TEST = 100;  // Максимальное количество заданий в тесте

    /**
     * Парсинг данных учеников с листа "Сбор информации"
     */
    public List<StudentResult> parseStudentData(Sheet dataSheet,
                                                Map<Integer, Integer> maxScores,
                                                String subject,
                                                String className) {
        List<StudentResult> results = new ArrayList<>();

        if (maxScores.isEmpty()) {
            log.error("Нет данных о максимальных баллов для предмета {}. Невозможно парсить учеников.", subject);
            return results;
        }

        log.debug("Парсинг учеников для предмета: {}, класс: {}", subject, className);
        log.debug("Всего строк в листе: {}", dataSheet.getLastRowNum() + 1);
        log.debug("Ожидаемое количество заданий: {}", maxScores.size());

        // 1. Определяем структуру колонок
        Row headerRow = dataSheet.getRow(HEADER_ROW_INDEX);
        if (headerRow == null) {
            log.error("Не найдена строка с заголовками заданий (строка {})", HEADER_ROW_INDEX + 1);
            return results;
        }

        // 2. Анализируем структуру колонок
        ColumnStructure columnStructure = analyzeColumnStructure(headerRow, maxScores.size());
        if (!columnStructure.isValid()) {
            log.error("Не удалось определить структуру колонок для файла");
            return results;
        }

        log.info("Структура файла: задания [{}-{}] ({} заданий), Итог в колонке {}",
                columnStructure.firstTaskColumn,
                columnStructure.lastTaskColumn,
                columnStructure.getTaskCount(),
                columnStructure.totalScoreColumn);

        // 3. Парсим студентов
        for (int rowIdx = FIRST_STUDENT_ROW; rowIdx <= dataSheet.getLastRowNum(); rowIdx++) {
            Row row = dataSheet.getRow(rowIdx);
            if (row == null) {
                continue;
            }

            StudentResult result = parseStudentRow(row, maxScores, columnStructure,
                    subject, className, rowIdx);

            if (result != null) {
                results.add(result);
            }
        }

        log.info("Найдено {} учеников для предмета {}", results.size(), subject);
        return results;
    }

    /**
     * Анализ структуры колонок файла
     */
    private ColumnStructure analyzeColumnStructure(Row headerRow, int expectedTaskCount) {
        ColumnStructure structure = new ColumnStructure();

        // 1. Ищем колонку "Итог" или "Итого" (последняя колонка с данными)
        for (int colIdx = 0; colIdx < headerRow.getLastCellNum(); colIdx++) {
            Cell cell = headerRow.getCell(colIdx);
            String cellValue = ExcelParser.getCellValueAsString(cell);

            if (cellValue != null &&
                    (cellValue.equalsIgnoreCase("Итог") ||
                            cellValue.equalsIgnoreCase("Итого"))) {
                structure.totalScoreColumn = colIdx;
                log.debug("Найдена колонка 'Итог' в индексе {}", colIdx);
                break;
            }
        }

        // Если не нашли "Итог", используем последнюю колонку как итоговую
        if (structure.totalScoreColumn == -1) {
            structure.totalScoreColumn = headerRow.getLastCellNum() - 1;
            log.warn("Не найдена колонка 'Итог', используем последнюю колонку {} как итоговую",
                    structure.totalScoreColumn);
        }

        // 2. Ищем начало заданий (первая колонка с номером задания)
        // Начинаем поиск с колонки D (после варианта)
        for (int colIdx = COL_VARIANT + 1; colIdx < structure.totalScoreColumn; colIdx++) {
            Cell cell = headerRow.getCell(colIdx);
            String cellValue = ExcelParser.getCellValueAsString(cell);

            if (isTaskNumber(cellValue)) {
                structure.firstTaskColumn = colIdx;
                log.debug("Найдено первое задание №{} в колонке {}", cellValue, colIdx);
                break;
            }
        }

        if (structure.firstTaskColumn == -1) {
            log.error("Не удалось найти начало заданий");
            return structure;
        }

        // 3. Определяем последнюю колонку с заданиями
        structure.lastTaskColumn = structure.totalScoreColumn - 1;

        // 4. Рассчитываем фактическое количество заданий
        int actualTaskCount = structure.lastTaskColumn - structure.firstTaskColumn + 1;
        structure.detectedTaskCount = actualTaskCount;

        log.info("Обнаружено {} заданий (ожидалось {})", actualTaskCount, expectedTaskCount);

        // 5. Проверяем соответствие ожидаемому количеству
        if (actualTaskCount != expectedTaskCount) {
            log.warn("⚠️ ВНИМАНИЕ: Количество заданий в файле ({}) не совпадает с ожидаемым ({})",
                    actualTaskCount, expectedTaskCount);

            // Если разница небольшая, можем работать с тем, что есть
            if (Math.abs(actualTaskCount - expectedTaskCount) <= 5) {
                log.warn("Используем фактическое количество заданий: {}", actualTaskCount);
            } else {
                log.error("Большое расхождение в количестве заданий! Файл может быть поврежден.");
            }
        }

        return structure;
    }

    /**
     * Проверяет, является ли строка номером задания
     */
    private boolean isTaskNumber(String value) {
        if (value == null || value.trim().isEmpty()) {
            return false;
        }

        value = value.trim();

        // Игнорируем текстовые заголовки
        String lowerValue = value.toLowerCase();
        if (lowerValue.contains("баллы") ||
                lowerValue.contains("задания") ||
                lowerValue.contains("ответы") ||
                lowerValue.contains("задание")) {
            return false;
        }

        // Пробуем распарсить как число
        try {
            // Убираем возможные точки в конце (например, "1.")
            if (value.endsWith(".")) {
                value = value.substring(0, value.length() - 1);
            }

            double num = Double.parseDouble(value);
            // Номер задания должен быть положительным числом
            return num > 0 && num <= 100; // Максимум 100 заданий
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * Парсинг одной строки с данными ученика
     */
    private StudentResult parseStudentRow(Row row,
                                          Map<Integer, Integer> maxScores,
                                          ColumnStructure columnStructure,
                                          String subject,
                                          String className,
                                          int rowIndex) {
        // 1. ФИО (колонка B)
        String fio = ExcelParser.getCellValueAsString(row.getCell(COL_FIO));
        if (fio == null || fio.trim().isEmpty() || fio.equals("[ПУСТО]")) {
            log.debug("Строка {}: ФИО пустое или '[ПУСТО]'", rowIndex + 1);
            return null;
        }
        fio = fio.trim();

        // 2. Присутствие (колонка C)
        String presence = ExcelParser.getCellValueAsString(row.getCell(COL_PRESENCE));
        if (presence == null || presence.trim().isEmpty()) {
            presence = "Не указано";
        } else {
            presence = presence.trim();
        }

        // 3. Вариант (колонка D)
        String variant = ExcelParser.getCellValueAsString(row.getCell(COL_VARIANT));
        if (variant == null) {
            variant = "";
        } else {
            variant = variant.trim();
        }

        // 4. Парсинг баллов за задания
        Map<Integer, Integer> taskScores = parseTaskScores(row, maxScores, columnStructure, rowIndex);

        // 5. Итоговый балл из колонки Итог
        Integer totalScore = parseTotalScore(row, columnStructure.totalScoreColumn, rowIndex);

        // 6. Создание результата
        StudentResult result = new StudentResult();
        result.setFio(fio);
        result.setPresence(presence);
        result.setVariant(variant);
        result.setSubject(subject);
        result.setClassName(className);
        result.setTaskScores(taskScores);

        // 7. Устанавливаем итоговый балл (приоритет - колонка Итог)
        if (totalScore != null) {
            result.setTotalScore(totalScore);

            // Валидация: сравниваем с суммой баллов за задания
            int calculatedTotal = JsonScoreUtils.calculateTotalScore(taskScores);
            if (!totalScore.equals(calculatedTotal)) {
                double difference = Math.abs(totalScore - calculatedTotal);
                log.warn("Строка {}: Студент {} - расхождение баллов: Итог={}, сумма={} (разница={})",
                        rowIndex + 1, fio, totalScore, calculatedTotal, difference);

                // Если разница небольшая (1-2 балла), вероятно округление
                if (difference <= 2) {
                    log.debug("Небольшое расхождение, вероятно округление");
                }
            }
        } else {
            // Если не удалось получить из колонки Итог, расчитываем
            result.setTotalScore(JsonScoreUtils.calculateTotalScore(taskScores));
            log.debug("Строка {}: Студент {} - итоговый балл расчитан", rowIndex + 1, fio);
        }

        return result;
    }

    /**
     * Парсинг баллов за задания
     */
    private Map<Integer, Integer> parseTaskScores(Row row,
                                                  Map<Integer, Integer> maxScores,
                                                  ColumnStructure columnStructure,
                                                  int rowIndex) {
        Map<Integer, Integer> scores = new HashMap<>();

        int taskNumber = 1;
        int parsedColumns = 0;

        // Парсим все задания от firstTaskColumn до lastTaskColumn
        for (int colIdx = columnStructure.firstTaskColumn;
             colIdx <= columnStructure.lastTaskColumn;
             colIdx++) {

            Cell cell = row.getCell(colIdx);
            Integer score = ExcelParser.getCellValueAsInteger(cell);

            if (score == null) {
                score = 0; // Если ячейка пустая
            }

            // Валидация: не превышает ли балл максимум (если информация есть)
            if (maxScores.containsKey(taskNumber)) {
                Integer maxForTask = maxScores.get(taskNumber);
                if (maxForTask != null && score > maxForTask) {
                    log.debug("Строка {}: Задание №{}: балл {} > максимум {}. Корректируем.",
                            rowIndex + 1, taskNumber, score, maxForTask);
                    score = maxForTask;
                }
            }

            scores.put(taskNumber, score);
            taskNumber++;
            parsedColumns++;

            // Защита от слишком большого количества заданий
            if (taskNumber > MAX_NUMBER_TEST) {
                log.error("Превышено максимальное количество заданий MAX_NUMBER_TEST = (MAX_NUMBER_TEST)");
                break;
            }
        }

        log.debug("Строка {}: Распарсено {} заданий", rowIndex + 1, parsedColumns);

        // Проверяем, что распарсили все ожидаемые задания
        if (parsedColumns != columnStructure.detectedTaskCount) {
            log.warn("Строка {}: Распарсено {} заданий, но ожидалось {}",
                    rowIndex + 1, parsedColumns, columnStructure.detectedTaskCount);
        }

        return scores;
    }

    /**
     * Парсинг итогового балла из колонки Итог
     */
    private Integer parseTotalScore(Row row, int totalScoreColumn, int rowIndex) {
        if (totalScoreColumn < 0 || totalScoreColumn >= row.getLastCellNum()) {
            return null;
        }

        Cell totalCell = row.getCell(totalScoreColumn);
        if (totalCell == null) {
            return null;
        }

        return ExcelParser.getCellValueAsInteger(totalCell);
    }

    /**
     * Класс для хранения информации о структуре колонок
     */
    private static class ColumnStructure {
        int firstTaskColumn = -1;      // Первая колонка с заданиями
        int lastTaskColumn = -1;       // Последняя колонка с заданиями
        int totalScoreColumn = -1;     // Колонка с итоговым баллом
        int detectedTaskCount = 0;     // Фактическое количество заданий

        boolean isValid() {
            return firstTaskColumn > 0 &&
                    lastTaskColumn >= firstTaskColumn &&
                    totalScoreColumn > lastTaskColumn &&
                    detectedTaskCount > 0;
        }

        int getTaskCount() {
            return detectedTaskCount;
        }
    }
}