package org.school.analysis.util;

import java.time.format.DateTimeFormatter;

public final class DateTimeFormatters {

    // Приватный конструктор для предотвращения инстанцирования
    private DateTimeFormatters() {
        throw new AssertionError("Utility class should not be instantiated");
    }

    // Формат для отображения дат в UI
    public static final DateTimeFormatter DISPLAY_DATE =
            DateTimeFormatter.ofPattern("dd.MM.yyyy");

    // Формат для имен файлов с временем
    public static final DateTimeFormatter FILE_WITH_TIME =
            DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

    // Формат для имен файлов без времени
    public static final DateTimeFormatter FILE_SIMPLE =
            DateTimeFormatter.ofPattern("yyyyMMdd");

    // Формат для логов
    public static final DateTimeFormatter LOG_TIMESTAMP =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    // Формат для базы данных
    public static final DateTimeFormatter SQL_DATE =
            DateTimeFormatter.ofPattern("yyyy-MM-dd");
}