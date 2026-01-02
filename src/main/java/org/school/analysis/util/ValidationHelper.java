package org.school.analysis.util;

import org.school.analysis.model.StudentResult;
import org.school.analysis.model.TestMetadata;

import java.util.Map;
import java.util.regex.Pattern;

/**
 * Утилиты для валидации данных
 */
public class ValidationHelper {

    // Паттерны для валидации
    private static final Pattern FIO_PATTERN =
            Pattern.compile("^[А-ЯЁ][а-яё]+\\s[А-ЯЁ][а-яё]+(\\s[А-ЯЁ][а-яё]+)?$");
    private static final Pattern CLASS_NAME_PATTERN =
            Pattern.compile("^\\d{1,2}[А-Яа-я]?$");
    private static final Pattern SUBJECT_PATTERN =
            Pattern.compile("^[А-Яа-яЁё\\s\\-]{3,50}$");

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
     * Валидация названия класса
     */
    public static boolean isValidClassName(String className) {
        if (className == null || className.trim().isEmpty()) {
            return false;
        }

        return CLASS_NAME_PATTERN.matcher(className.trim()).matches();
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
        if (!isValidClassName(student.getClassName())) {
            result.addWarning("Некорректный формат класса: " + student.getClassName());
        }

        // Валидация предмета
        if (!isValidSubject(student.getSubject())) {
            result.addWarning("Некорректный предмет: " + student.getSubject());
        }

        // Валидация баллов
        if (student.getTaskScores() != null && maxScores != null) {
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

        // Валидация количества заданий
        if (metadata.getTaskCount() <= 0 || metadata.getTaskCount() > 100) {
            result.addError("Некорректное количество заданий: " + metadata.getTaskCount());
        }

        // Валидация максимальных баллов
        if (metadata.getMaxScores() == null || metadata.getMaxScores().isEmpty()) {
            result.addError("Отсутствуют максимальные баллы");
        } else {
            for (var entry : metadata.getMaxScores().entrySet()) {
                if (entry.getKey() <= 0 || entry.getKey() > metadata.getTaskCount()) {
                    result.addError("Некорректный номер задания: " + entry.getKey());
                }
                if (entry.getValue() <= 0 || entry.getValue() > 100) {
                    result.addError("Некорректный максимальный балл: " + entry.getValue());
                }
            }
        }

        return result;
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

        public java.util.List<String> getErrors() {
            return new java.util.ArrayList<>(errors);
        }

        public java.util.List<String> getWarnings() {
            return new java.util.ArrayList<>(warnings);
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
}