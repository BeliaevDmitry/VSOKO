package org.school.analysis.config;

import java.util.List;

public class AppConfig {
    // ========== КОНФИГУРАЦИЯ ==========

    public static final List<String> SCHOOLS = List.of(
            "ГБОУ №7",
            "ГБОУ №1811"
    );


    public static final List<String> ALL_ACADEMIC_YEAR = List.of(
            "2025-2026",
            "2026-2027"
    );

    // Папка с исходными файлами (учителя кладут сюда Excel файлы)
    public static final String INPUT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\{школа}\\ВСОКО\\Работы\\На разбор";

        // Шаблон для папок с обработанными отчетами
    public static final String REPORTS_BASE_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\{школа}\\ВСОКО\\Работы\\{предмет}\\Отчёты";
    public static final String REPORTS_ANALISIS_BASE_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\{школа}\\ВСОКО\\Работы\\{предмет}\\Анализ";

    // Папка для итоговых отчетов (сводная статистика)
    public static final String FINAL_REPORT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\{школа}\\ВСОКО\\Работы";

    // Папка для сохранения статистики отчетов (сводная статистика)
    public static final String STATISTIK_REPORT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\отчёты программ";



    public static final String INPUT_TEACHER_NAME =
            "C:\\Users\\dimah\\Yandex.Disk\\{школа}\\Учителя.xlsx";

    // Настройки для импорта учителей
    public static final boolean AUTO_IMPORT_TEACHERS = true;
    public static final int TEACHER_IMPORT_CHECK_INTERVAL = 300000; // 5 минут в миллисекундах

    // Добавьте это для контроля размера пакета
    public static final int BATCH_SIZE = 100;
}
