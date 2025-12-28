package org.school.analysis;

import org.school.analysis.repository.StudentResultRepository;
import org.school.analysis.repository.impl.InMemoryStudentRepository;
import org.school.analysis.service.*;
import org.school.analysis.service.impl.*;

/**
 * Главный класс приложения
 * Содержит все конфигурационные константы
 */
public class Main {

    // ========== КОНФИГУРАЦИЯ ==========

    // Папка с исходными файлами (учителя кладут сюда Excel файлы)
    private static final String INPUT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\На разбор";

    // Шаблон для папок с обработанными отчетами
    // {предмет} будет заменен на название предмета
    private static final String REPORTS_BASE_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\{предмет}\\Отчёты";

    // Папка для итоговых отчетов (сводная статистика)
    private static final String FINAL_REPORT_FOLDER =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО";

    // Максимальное количество учеников в классе
    private static final int MAX_STUDENTS_PER_CLASS = 34;

    // Максимальное количество заданий в тесте
    private static final int MAX_TASKS_PER_TEST = 30;

    // ========== ТОЧКА ВХОДА ==========

    public static void main(String[] args) {
        System.out.println("=== ЗАПУСК СИСТЕМЫ ОБРАБОТКИ ОТЧЁТОВ ВСОКО ===");
        System.out.println("Конфигурация:");
        System.out.println("  Папка с исходными файлами: " + INPUT_FOLDER);
        System.out.println("  Шаблон для отчетов: " + REPORTS_BASE_FOLDER);
        System.out.println("  Папка для итогов: " + FINAL_REPORT_FOLDER);
        System.out.println();

        try {
            // 1. Инициализация всех сервисов
            ReportFileFinderService fileFinder = new ReportFileFinderServiceImpl();
            ReportParserService parser = new ReportParserServiceImpl();
            StudentResultRepository repository = new InMemoryStudentRepository();
            FileOrganizerService fileOrganizer = new FileOrganizerServiceImpl(REPORTS_BASE_FOLDER);

            // 2. Создание главного сервиса
            ReportProcessorService processor = new ReportProcessorServiceImpl(
                    fileFinder,
                    parser,
                    repository,
                    fileOrganizer
            );

            // 3. Запуск обработки
            var summary = processor.processAll(INPUT_FOLDER);

            // 4. Вывод результатов
            printSummary(summary);

        } catch (Exception e) {
            System.err.println("❌ Критическая ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void printSummary(ReportProcessorService.ProcessingSummary summary) {
        System.out.println("\n" + "=".repeat(50));
        System.out.println("ИТОГИ ОБРАБОТКИ:");
        System.out.println("=".repeat(50));
        System.out.println("📁 Всего найдено файлов: " + summary.getTotalFilesFound());
        System.out.println("✅ Успешно распарсено: " + summary.getSuccessfullyParsed());
        System.out.println("💾 Сохранено в БД: " + summary.getSuccessfullySaved());
        System.out.println("📂 Перемещено файлов: " + summary.getSuccessfullyMoved());
        System.out.println("👨‍🎓 Обработано учеников: " + summary.getTotalStudentsProcessed());

        if (!summary.getFailedFiles().isEmpty()) {
            System.out.println("\n⚠️ ФАЙЛЫ С ОШИБКАМИ:");
            for (var file : summary.getFailedFiles()) {
                System.out.println("  • " + file.getFile().getName() +
                        " - " + file.getErrorMessage());
            }
        }

        System.out.println("\n" + "=".repeat(50));
    }
}