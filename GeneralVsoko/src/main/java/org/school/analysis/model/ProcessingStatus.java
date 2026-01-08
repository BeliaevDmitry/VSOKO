package org.school.analysis.model;

public enum ProcessingStatus {
    PENDING("В ожидании"),
    PROCESSING("В обработке"),
    PARSED("Распарсен"),
    SAVED("Сохранен"),
    MOVED("Перемещен"),
    ERROR_PARSING("Ошибка парсинга"),
    ERROR_SAVING("Ошибка сохранения"),
    ERROR_MOVING("Ошибка перемещения"),
    ERROR_VALIDATION("Ошибка валидации"),
    DUPLICATE("Дубликат"),
    SKIPPED("Пропущен");

    private final String description;

    ProcessingStatus(String description) {
        this.description = description;
    }

    public String getDescription() {
        return description;
    }
}