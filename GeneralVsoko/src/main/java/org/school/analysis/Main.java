package org.school.analysis;

import org.school.analysis.service.GeneralService;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;

import static org.school.analysis.config.AppConfig.*;

@SpringBootApplication
public class Main {
    public static void main(String[] args) {
        ApplicationContext context = SpringApplication.run(Main.class, args);
        GeneralService processor = context.getBean(GeneralService.class);

        System.out.println("=".repeat(80));
        System.out.println("🚀 ЗАПУСК СИСТЕМЫ ОБРАБОТКИ ОТЧЁТОВ ВСОКО");
        System.out.println("=".repeat(80));
        System.out.println("Конфигурация:");
        System.out.println("  📁 Шаблон с исходными файлами: " + INPUT_FOLDER);
        System.out.println("  📊 Шаблон для отчетов: " + REPORTS_BASE_FOLDER);
        System.out.println("  📂 Папка для итогов: " + FINAL_REPORT_FOLDER);
        System.out.println("  🏫 Школы для обработки: " + String.join(", ", SCHOOLS));
        System.out.println("  📅 Текущий учебный год: " + ALL_ACADEMIC_YEAR.get(0));
        System.out.println("=".repeat(80));
        System.out.println();

        try {
            processor.processAll();
        } catch (Exception e) {
            System.err.println("❌ Критическая ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }
}