package org.school.analysis;

import org.school.analysis.model.ProcessingSummary;
import org.school.analysis.service.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;

import static org.school.analysis.config.AppConfig.*;

@SpringBootApplication
public class Main {
    public static void main(String[] args) {
        // Запускаем Spring контекст
        ApplicationContext context = SpringApplication.run(Main.class, args);

        // Получаем главный сервис из контекста
        ReportProcessorService processor = context.getBean(ReportProcessorService.class);


        System.out.println("=== ЗАПУСК СИСТЕМЫ ОБРАБОТКИ ОТЧЁТОВ ВСОКО ===");
        System.out.println("Конфигурация:");
        System.out.println("  Папка с исходными файлами: " + INPUT_FOLDER);
        System.out.println("  Шаблон для отчетов: " + REPORTS_BASE_FOLDER);
        System.out.println("  Папка для итогов: " + FINAL_REPORT_FOLDER);
        System.out.println();

        try {
            // 3. Запуск обработки
            var summary = processor.processAll(INPUT_FOLDER);

            // 4. Вывод результатов
            printSummary(summary);

        } catch (Exception e) {
            System.err.println("❌ Критическая ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void printSummary(ProcessingSummary summary) {
        System.out.println("\n" + "=".repeat(50));
        System.out.println("ИТОГИ ОБРАБОТКИ:");
        System.out.println("=".repeat(50));
        System.out.println("📁 Всего найдено файлов: " + summary.getTotalFilesFound());
        System.out.println("✅ Успешно распарсено: " + summary.getSuccessfullyParsed());
        System.out.println("💾 Сохранено в БД: " + summary.getSuccessfullySaved());
        System.out.println("📂 Перемещено файлов: " + summary.getSuccessfullyMoved());
        System.out.println("\n" + "=".repeat(50));
    }
}