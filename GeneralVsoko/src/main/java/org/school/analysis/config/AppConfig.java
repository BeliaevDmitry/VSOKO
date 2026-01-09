package org.school.analysis.config;

public class AppConfig {
    // ========== КОНФИГУРАЦИЯ ==========

    // Папка с исходными файлами (учителя кладут сюда Excel файлы)
    public static final String INPUT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\На разбор";

    // Шаблон для папок с обработанными отчетами
    // {предмет} будет заменен на название предмета
    public static final String REPORTS_BASE_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\{предмет}\\Отчёты";


    public static final String REPORTS_Analisis_BASE_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\{предмет}\\Анализ";
    // Папка для итоговых отчетов (сводная статистика)
    public static final String FINAL_REPORT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы";

    // Максимальное количество учеников в классе
    public static final int MAX_STUDENTS_PER_CLASS = 34;

    // Максимальное количество заданий в тесте
    public static final int MAX_TASKS_PER_TEST = 30;

    // Добавьте это для контроля размера пакета
    public static final int BATCH_SIZE = 100;
}
