package org.school.analysis.model;

public enum ProcessingStatus {
    PENDING,      // Ожидает обработки
    PARSING,      // В процессе парсинга
    PARSED,       // Успешно распарсен
    SAVED,        // Сохранен в БД
    MOVED,        // Перемещен в архив
    ERROR_PARSING, // Ошибка парсинга
    ERROR_SAVING,  // Ошибка сохранения
    ERROR_MOVING   // Ошибка перемещения
}