package org.school.analysis.exception;

/**
 * Исключение для ошибок валидации данных
 */
public class ValidationException extends RuntimeException {

    public ValidationException(String message) {
        super(message);
    }

    public ValidationException(String message, Throwable cause) {
        super(message, cause);
    }

    public ValidationException(String message, String fileName, int row) {
        super(String.format("Ошибка в файле '%s', строка %d: %s", fileName, row, message));
    }
}