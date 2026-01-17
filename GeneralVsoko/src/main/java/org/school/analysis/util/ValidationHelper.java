package org.school.analysis.util;

import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.model.TestMetadata;
import org.school.analysis.service.TeacherService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;
import java.util.regex.Pattern;

/**
 * Утилиты для валидации данных
 */
public class ValidationHelper {

    private static final Logger log = LoggerFactory.getLogger(ValidationHelper.class);

    // Паттерны для валидации
    private static final Pattern FIO_PATTERN =
            Pattern.compile("^[А-ЯЁ][а-яё]+\\s[А-ЯЁ][а-яё]+(\\s[А-ЯЁ][а-яё]+)?$");
    private static final Pattern CLASS_NAME_PATTERN =
            Pattern.compile("^(1[0-1]|[1-9])[-][А-Яа-я]$");
    private static final Pattern SUBJECT_PATTERN =
            Pattern.compile("^[А-Яа-яЁё\\s\\-()]{3,100}$");

    /**
     * Валидация ФИО
     */
    public static boolean isValidFio(String fio) {
        if (fio == null || fio.trim().isEmpty()) {
            return false;
        }

        return FIO_PATTERN.matcher(fio.trim()).matches();
    }

    /**
     * Валидация названия класса с поддержкой адресов и специальных названий
     */
    public static boolean isValidClassName(String className) {
        if (className == null || className.trim().isEmpty()) {
            return false;
        }

        String trimmed = className.trim();

        // 1. Стандартный формат класса: "11-А", "10-Б" и т.д.
        if (CLASS_NAME_PATTERN.matcher(trimmed).matches()) {
            return true;
        }

        // 2. Проверка на адреса (начинается с "ул.", "корпус", "адрес" и т.д.)
        String lowerTrimmed = trimmed.toLowerCase();
        if (lowerTrimmed.startsWith("ул.") ||
                lowerTrimmed.startsWith("улица") ||
                lowerTrimmed.startsWith("корпус") ||
                lowerTrimmed.startsWith("адрес") ||
                lowerTrimmed.startsWith("здание") ||
                lowerTrimmed.contains("марии") ||
                lowerTrimmed.contains("кравченко")) {
            return true;
        }

        // 3. Специальные значения
        if (trimmed.equals("Без группы") ||
                trimmed.equals("Без класса") ||
                trimmed.equals("Без адреса") ||
                trimmed.equals("Неизвестный класс")) {
            return true;
        }

        // 4. Форматы типа "11А", "10Б" (без дефиса)
        if (trimmed.matches("^(1[0-1]|[1-9])[А-Яа-я]$")) {
            return true;
        }

        // 5. Просто число класса "11", "10"
        if (trimmed.matches("^(1[0-1]|[1-9])$")) {
            return true;
        }

        // 6. Если содержит буквы и цифры (минимальная проверка)
        if (trimmed.matches(".*[0-9].*") && trimmed.matches(".*[А-Яа-яA-Za-z].*")) {
            return true;
        }

        // 7. Допустим любую строку длиной 2-50 символов
        // (последнее правило - максимально либеральное)
        return trimmed.length() >= 2 && trimmed.length() <= 50;
    }

    /**
     * Валидация названия предмета
     */
    public static boolean isValidSubject(String subject) {
        if (subject == null || subject.trim().isEmpty()) {
            return false;
        }

        return SUBJECT_PATTERN.matcher(subject.trim()).matches();
    }

    /**
     * Валидация баллов (должны быть в пределах 0-макс)
     */
    public static boolean isValidScore(int score, int maxScore) {
        return score >= 0 && score <= maxScore;
    }

    /**
     * Валидация результата ученика
     */
    public static ValidationResult validateStudentResult(StudentResult student, Map<Integer, Integer> maxScores) {
        ValidationResult result = new ValidationResult();

        if (student == null) {
            result.addError("Результат ученика не может быть null");
            return result;
        }

        // Валидация ФИО
        if (!isValidFio(student.getFio())) {
            result.addError("Некорректный формат ФИО: " + student.getFio());
        }

        // Валидация класса
        if (student.getClassName() != null && !isValidClassName(student.getClassName())) {
            result.addWarning("Некорректный формат класса: " + student.getClassName());
        }

        // Валидация предмета
        if (student.getSubject() != null && !isValidSubject(student.getSubject())) {
            result.addWarning("Некорректный предмет: " + student.getSubject());
        }

        // Валидация присутствия
        String presence = student.getPresence();
        if (presence == null || presence.trim().isEmpty()) {
            result.addError("Не указано присутствие");
        } else if (!isValidPresence(presence)) {
            result.addWarning("Некорректное значение присутствия: " + presence);
        }

        // Валидация баллов
        if (student.wasPresent() && student.getTaskScores() != null && maxScores != null) {
            for (var entry : student.getTaskScores().entrySet()) {
                Integer taskNum = entry.getKey();
                Integer score = entry.getValue();
                Integer maxScore = maxScores.get(taskNum);

                if (maxScore == null) {
                    result.addError("Отсутствует максимальный балл для задания " + taskNum);
                } else if (!isValidScore(score, maxScore)) {
                    result.addError(String.format(
                            "Некорректный балл для задания %d: %d (максимум %d)",
                            taskNum, score, maxScore
                    ));
                }
            }

            // Валидация общего балла (если он установлен)
            if (student.getTotalScore() != null) {
                int calculatedTotal = JsonScoreUtils.calculateTotalScore(student.getTaskScores());
                if (student.getTotalScore() != calculatedTotal) {
                    result.addWarning(String.format(
                            "Несоответствие общего балла: указано %d, рассчитано %d",
                            student.getTotalScore(), calculatedTotal
                    ));
                }
            }
        }

        return result;
    }

    /**
     * Валидация метаданных теста
     */
    public static ValidationResult validateTestMetadata(TestMetadata metadata) {
        ValidationResult result = new ValidationResult();

        if (metadata == null) {
            result.addError("Метаданные не могут быть null");
            return result;
        }

        // Валидация предмета
        if (!isValidSubject(metadata.getSubject())) {
            result.addError("Некорректный предмет: " + metadata.getSubject());
        }

        // Валидация класса
        if (!isValidClassName(metadata.getClassName())) {
            result.addError("Некорректный класс: " + metadata.getClassName());
        }

        // Валидация даты теста
        if (metadata.getTestDate() == null) {
            result.addError("Не указана дата теста");
        }

        // Валидация максимальных баллов
        if (metadata.getMaxScores() == null || metadata.getMaxScores().isEmpty()) {
            result.addError("Отсутствуют максимальные баллы");
        } else {
            for (var entry : metadata.getMaxScores().entrySet()) {
                Integer taskNum = entry.getKey();
                Integer maxScore = entry.getValue();

                if (taskNum <= 0 || taskNum > 100) {
                    result.addError("Некорректный номер задания: " + taskNum);
                }
                if (maxScore <= 0 || maxScore > 100) {
                    result.addError(String.format(
                            "Некорректный максимальный балл для задания %d: %d",
                            taskNum, maxScore
                    ));
                }
            }

            // Проверяем последовательность номеров заданий
            int expectedTaskNum = 1;
            for (Integer taskNum : metadata.getMaxScores().keySet().stream().sorted().toList()) {
                if (taskNum != expectedTaskNum) {
                    result.addWarning(String.format(
                            "Пропущен номер задания: ожидалось %d, найдено %d",
                            expectedTaskNum, taskNum
                    ));
                }
                expectedTaskNum++;
            }
        }

        return result;
    }

    /**
     * Валидация ReportFile - возвращает true/false и логирует ошибки
     */
    public static boolean validateReportFile(ReportFile reportFile) {
        if (reportFile == null) {
            log.error("ReportFile не может быть null");
            return false;
        }

        boolean isValid = true;
        String fileName = reportFile.getFileName();

        // Валидация файла
        if (reportFile.getFile() == null) {
            log.error("Файл {}: файл не может быть null", fileName);
            isValid = false;
        }

        // Валидация предмета
        if (!isValidSubject(reportFile.getSubject())) {
            log.error("Файл {}: некорректный предмет: {}", fileName, reportFile.getSubject());
            isValid = false;
        }

        // Валидация класса
        if (!isValidClassName(reportFile.getClassName())) {
            log.error("Файл {}: некорректный класс: {}", fileName, reportFile.getClassName());
            isValid = false;
        }

        // Валидация даты теста
        if (reportFile.getTestDate() == null) {
            log.error("Файл {}: не указана дата теста", fileName);
            isValid = false;
        }

        // Валидация максимальных баллов
        Map<Integer, Integer> maxScores = reportFile.getMaxScores();
        if (maxScores == null || maxScores.isEmpty()) {
            log.error("Файл {}: отсутствуют максимальные баллы", fileName);
            isValid = false;
        } else {
            // Проверяем согласованность taskCount и количества заданий
            int actualTaskCount = maxScores.size();
            if (reportFile.getTaskCount() != actualTaskCount) {
                log.warn("Файл {}: несоответствие количества заданий: указано {}, найдено {}",
                        fileName, reportFile.getTaskCount(), actualTaskCount);
                // Автоматически исправляем
                reportFile.setTaskCount(actualTaskCount);
            }

            // Валидация каждого задания
            for (var entry : maxScores.entrySet()) {
                Integer taskNum = entry.getKey();
                Integer maxScore = entry.getValue();

                if (taskNum <= 0 || taskNum > 100) {
                    log.error("Файл {}: некорректный номер задания: {}", fileName, taskNum);
                    isValid = false;
                }
                if (maxScore <= 0 || maxScore > 100) {
                    log.error("Файл {}: некорректный максимальный балл для задания {}: {}",
                            fileName, taskNum, maxScore);
                    isValid = false;
                }
            }
        }

        // Валидация количества студентов
        if (reportFile.getStudentCount() < 0) {
            log.error("Файл {}: некорректное количество студентов: {}",
                    fileName, reportFile.getStudentCount());
            isValid = false;
        }

        if (isValid) {
            log.debug("Файл {}: валидация пройдена успешно", fileName);
        } else {
            log.warn("Файл {}: валидация не пройдена", fileName);
        }

        return isValid;
    }

    /**
     * Расширенная валидация ReportFile с возвратом результата
     */
    public static ValidationResult validateReportFileDetailed(ReportFile reportFile) {
        ValidationResult result = new ValidationResult();
        String fileName = reportFile != null && reportFile.getFile() != null
                ? reportFile.getFileName()
                : "unknown";

        if (reportFile == null) {
            result.addError("ReportFile не может быть null");
            log.error("Файл unknown: ReportFile не может быть null");
            return result;
        }

        // Валидация файла
        if (reportFile.getFile() == null) {
            result.addError("Файл не может быть null");
            log.error("Файл {}: файл не может быть null", fileName);
        }

        // Валидация предмета
        if (!isValidSubject(reportFile.getSubject())) {
            String error = "Некорректный предмет: " + reportFile.getSubject();
            result.addError(error);
            log.error("Файл {}: {}", fileName, error);
        }

        // Валидация класса
        if (!isValidClassName(reportFile.getClassName())) {
            String error = "Некорректный класс: " + reportFile.getClassName();
            result.addError(error);
            log.error("Файл {}: {}", fileName, error);
        }

        // Валидация даты теста
        if (reportFile.getTestDate() == null) {
            String error = "Не указана дата теста";
            result.addError(error);
            log.error("Файл {}: {}", fileName, error);
        }

        // Валидация максимальных баллов
        Map<Integer, Integer> maxScores = reportFile.getMaxScores();
        if (maxScores == null || maxScores.isEmpty()) {
            String error = "Отсутствуют максимальные баллы";
            result.addError(error);
            log.error("Файл {}: {}", fileName, error);
        } else {
            // Проверяем согласованность taskCount и количества заданий
            int actualTaskCount = maxScores.size();
            if (reportFile.getTaskCount() != actualTaskCount) {
                String warning = String.format(
                        "Несоответствие количества заданий: указано %d, найдено %d",
                        reportFile.getTaskCount(), actualTaskCount
                );
                result.addWarning(warning);
                log.warn("Файл {}: {}", fileName, warning);
                // Автоматически исправляем
                reportFile.setTaskCount(actualTaskCount);
            }

            // Валидация каждого задания
            for (var entry : maxScores.entrySet()) {
                Integer taskNum = entry.getKey();
                Integer maxScore = entry.getValue();

                if (taskNum <= 0 || taskNum > 100) {
                    String error = "Некорректный номер задания: " + taskNum;
                    result.addError(error);
                    log.error("Файл {}: {}", fileName, error);
                }
                if (maxScore <= 0 || maxScore > 100) {
                    String error = String.format(
                            "Некорректный максимальный балл для задания %d: %d",
                            taskNum, maxScore
                    );
                    result.addError(error);
                    log.error("Файл {}: {}", fileName, error);
                }
            }
        }

        // Валидация количества студентов
        if (reportFile.getStudentCount() < 0) {
            String error = "Некорректное количество студентов: " + reportFile.getStudentCount();
            result.addError(error);
            log.error("Файл {}: {}", fileName, error);
        }

        if (result.isValid()) {
            log.debug("Файл {}: валидация пройдена успешно", fileName);
        }

        return result;
    }

    /**
     * Проверка корректности значения присутствия
     */
    private static boolean isValidPresence(String presence) {
        if (presence == null) {
            return false;
        }
        String normalized = presence.trim().toLowerCase();
        return normalized.equals("был") ||
                normalized.equals("была") ||
                normalized.equals("отсутствовал") ||
                normalized.equals("отсутствовала");
    }

    /**
     * Проверка, является ли строка числом
     */
    public static boolean isNumeric(String str) {
        if (str == null || str.trim().isEmpty()) {
            return false;
        }

        try {
            Double.parseDouble(str.replace(",", "."));
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * Проверка, является ли строка целым числом
     */
    public static boolean isInteger(String str) {
        if (str == null || str.trim().isEmpty()) {
            return false;
        }

        try {
            Integer.parseInt(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * Проверка email
     */
    public static boolean isValidEmail(String email) {
        if (email == null || email.trim().isEmpty()) {
            return false;
        }

        String emailPattern = "^[A-Za-z0-9+_.-]+@(.+)$";
        return Pattern.compile(emailPattern).matcher(email).matches();
    }

    /**
     * Валидация JSON строки с баллами
     */
    public static boolean isValidScoreJson(String json) {
        if (json == null || json.trim().isEmpty() || json.trim().equals("{}")) {
            return true; // Пустой JSON допустим для отсутствующих студентов
        }

        try {
            Map<Integer, Integer> scores = JsonScoreUtils.jsonToMap(json);
            if (scores == null) {
                return false;
            }
            // Проверяем, что все значения корректны
            for (var entry : scores.entrySet()) {
                if (entry.getKey() <= 0 || entry.getValue() < 0) {
                    return false;
                }
            }
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * Результат валидации
     */
    public static class ValidationResult {
        private boolean valid = true;
        private final java.util.List<String> errors = new java.util.ArrayList<>();
        private final java.util.List<String> warnings = new java.util.ArrayList<>();

        public void addError(String error) {
            errors.add(error);
            valid = false;
        }

        public void addWarning(String warning) {
            warnings.add(warning);
        }

        public boolean isValid() {
            return valid && errors.isEmpty();
        }

        public boolean hasErrors() {
            return !errors.isEmpty();
        }

        public boolean hasWarnings() {
            return !warnings.isEmpty();
        }

        public java.util.List<String> getErrors() {
            return new java.util.ArrayList<>(errors);
        }

        public java.util.List<String> getWarnings() {
            return new java.util.ArrayList<>(warnings);
        }

        public String getErrorsAsString() {
            return String.join("; ", errors);
        }

        public String getWarningsAsString() {
            return String.join("; ", warnings);
        }

        @Override
        public String toString() {
            StringBuilder sb = new StringBuilder();
            if (!errors.isEmpty()) {
                sb.append("Ошибки:\n");
                for (String error : errors) {
                    sb.append("  - ").append(error).append("\n");
                }
            }
            if (!warnings.isEmpty()) {
                sb.append("Предупреждения:\n");
                for (String warning : warnings) {
                    sb.append("  - ").append(warning).append("\n");
                }
            }
            return sb.toString();
        }
    }

    /**
     * Валидация ФИО учителя с проверкой в базе данных
     */
    public static ValidationResult validateTeacher(String teacherName, TeacherService teacherService) {
        ValidationResult result = new ValidationResult();

        if (teacherName == null || teacherName.trim().isEmpty()) {
            result.addError("ФИО учителя не может быть пустым");
            return result;
        }

        String trimmedName = teacherName.trim();

        // Базовые проверки формата
        if (!isValidTeacherNameFormat(trimmedName)) {
            result.addError("Некорректный формат ФИО учителя: " + trimmedName);
        }

        // Проверка в базе данных (если передан TeacherService)
        if (teacherService != null && !teacherService.isTeacherValid(trimmedName)) {
            result.addError("Учитель '" + trimmedName + "' не найден в базе данных");
        }

        return result;
    }

    /**
     * Валидация формата ФИО учителя (базовая проверка)
     */
    public static boolean isValidTeacherNameFormat(String teacherName) {
        if (teacherName == null || teacherName.trim().isEmpty()) {
            return false;
        }

        String trimmed = teacherName.trim();

        // 1. Допустимые форматы:
        // - Полное ФИО: "Иванов Иван Иванович"
        // - Сокращенное: "Иванов И.И."
        // - Только фамилия и инициалы: "Иванов И.И"
        // - С ошибками: "Иванов Иван Ивановыч" (проверяется fuzzy search в TeacherService)

        // 2. Минимальная длина
        if (trimmed.length() < 3) {
            return false;
        }

        // 3. Должны быть русские буквы
        if (!trimmed.matches(".*[А-Яа-яЁё].*")) {
            return false;
        }

        // 4. Не должен содержать запрещенных символов
        if (trimmed.matches(".*[0-9!@#$%^&*()_+=<>?/\\\\|].*")) {
            return false;
        }

        return true;
    }

    /**
     * Валидация Teacher из метаданных
     */
    public static ValidationResult validateTeacherInMetadata(TestMetadata metadata, TeacherService teacherService) {
        ValidationResult result = new ValidationResult();

        if (metadata == null) {
            result.addError("Метаданные не могут быть null");
            return result;
        }

        String teacherName = metadata.getTeacher();
        if (teacherName == null || teacherName.trim().isEmpty()) {
            result.addError("Не указан учитель");
            return result;
        }

        // Используем общую валидацию учителя
        ValidationResult teacherValidation = validateTeacher(teacherName, teacherService);
        if (teacherValidation.hasErrors()) {
            result.addError("Ошибка валидации учителя: " + teacherValidation.getErrorsAsString());
        }

        return result;
    }

    /**
     * Комплексная валидация ReportFile с проверкой учителя
     */
    public static ValidationResult validateReportFileWithTeacher(ReportFile reportFile, TeacherService teacherService) {
        ValidationResult result = validateReportFileDetailed(reportFile);

        // Дополнительно проверяем учителя
        if (reportFile != null && reportFile.getTeacher() != null) {
            ValidationResult teacherResult = validateTeacher(reportFile.getTeacher(), teacherService);
            if (teacherResult.hasErrors()) {
                result.addError("Ошибка валидации учителя: " + teacherResult.getErrorsAsString());
            }
            if (teacherResult.hasWarnings()) {
                result.addWarning("Предупреждение по учителю: " + teacherResult.getWarningsAsString());
            }
        }

        return result;
    }
}