package org.school.analysis.util;

import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.nio.file.Paths;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

import static org.school.analysis.config.AppConfig.STATISTIK_REPORT_FOLDER;

@Slf4j
public class PerformanceTracker {

    private static final Map<String, SchoolProcessingMetrics> schoolMetrics = new ConcurrentHashMap<>();
    private static LocalDateTime programStartTime;

    @Data
    @Builder
    public static class SchoolProcessingMetrics {
        private String schoolName;
        private LocalDateTime startTime;
        private LocalDateTime endTime;
        private Duration totalTime;
        private Duration fileFindingTime;
        private Duration fileProcessingTime;
        private Duration reportGenerationTime;
        private int filesFound;
        private int filesProcessed;
        private int reportsGenerated;

        public String getFormattedDuration() {
            if (totalTime == null) return "-";
            long seconds = totalTime.getSeconds();
            long millis = totalTime.toMillisPart();
            return String.format("%d.%03d сек", seconds, millis);
        }

        public double getSuccessRate() {
            return filesFound > 0 ? (filesProcessed * 100.0) / filesFound : 0;
        }
    }

    /**
     * Начать отсчет общего времени программы
     */
    public static void startProgram() {
        programStartTime = LocalDateTime.now();
        schoolMetrics.clear();
        log.info("🚀 Начало выполнения программы в {}",
                programStartTime.format(DateTimeFormatter.ofPattern("HH:mm:ss")));
    }

    /**
     * Начать отсчет времени обработки школы
     */
    public static SchoolProcessingMetrics startSchoolProcessing(String schoolName) {
        log.info("🎯 Начало обработки школы: {}", schoolName);
        return SchoolProcessingMetrics.builder()
                .schoolName(schoolName)
                .startTime(LocalDateTime.now())
                .build();
    }

    /**
     * Завершить обработку школы
     */
    public static void finishSchoolProcessing(SchoolProcessingMetrics metrics,
                                              int filesFound,
                                              int filesProcessed,
                                              int reportsGenerated) {
        if (metrics == null) return;

        metrics.setEndTime(LocalDateTime.now());
        metrics.setTotalTime(Duration.between(metrics.getStartTime(), metrics.getEndTime()));
        metrics.setFilesFound(filesFound);
        metrics.setFilesProcessed(filesProcessed);
        metrics.setReportsGenerated(reportsGenerated);

        schoolMetrics.put(metrics.getSchoolName(), metrics);

        log.info("✅ Завершена обработка школы {} за {}",
                metrics.getSchoolName(), metrics.getFormattedDuration());
    }

    /**
     * Записать время для конкретной фазы обработки
     */
    public static void recordPhaseTime(String schoolName, String phase, Duration duration) {
        SchoolProcessingMetrics metrics = schoolMetrics.get(schoolName);
        if (metrics != null && duration != null) {
            switch (phase) {
                case "fileFinding":
                    metrics.setFileFindingTime(duration);
                    break;
                case "fileProcessing":
                    metrics.setFileProcessingTime(duration);
                    break;
                case "reportGeneration":
                    metrics.setReportGenerationTime(duration);
                    break;
            }
        }
    }

    /**
     * Получить статистику по всем школам
     */
    public static String getSchoolsStatistics() {
        if (schoolMetrics.isEmpty()) {
            return "Нет данных по школам";
        }

        StringBuilder sb = new StringBuilder();
        sb.append("\n📊 СТАТИСТИКА ПО ШКОЛАМ\n");
        sb.append("=".repeat(100)).append("\n");
        sb.append(String.format("%-20s | %12s | %8s | %12s | %10s | %8s | %15s | %15s | %15s\n",
                "Школа", "Общее время", "Файлов", "Обработано", "Отчётов", "Эффектив.",
                "Поиск файлов", "Обработка", "Генерация отч."));
        sb.append("-".repeat(100)).append("\n");

        int totalFiles = 0;
        int totalProcessed = 0;
        int totalReports = 0;
        Duration totalTime = Duration.ZERO;

        List<SchoolProcessingMetrics> sortedMetrics = new ArrayList<>(schoolMetrics.values());
        sortedMetrics.sort(Comparator.comparing(SchoolProcessingMetrics::getTotalTime).reversed());

        for (SchoolProcessingMetrics metrics : sortedMetrics) {
            sb.append(String.format("%-20s | %12s | %8d | %12d | %10d | %7.1f%% | %15s | %15s | %15s\n",
                    metrics.getSchoolName(),
                    metrics.getFormattedDuration(),
                    metrics.getFilesFound(),
                    metrics.getFilesProcessed(),
                    metrics.getReportsGenerated(),
                    metrics.getSuccessRate(),
                    formatDuration(metrics.getFileFindingTime()),
                    formatDuration(metrics.getFileProcessingTime()),
                    formatDuration(metrics.getReportGenerationTime())));

            totalFiles += metrics.getFilesFound();
            totalProcessed += metrics.getFilesProcessed();
            totalReports += metrics.getReportsGenerated();
            totalTime = totalTime.plus(metrics.getTotalTime());
        }

        sb.append("-".repeat(100)).append("\n");
        sb.append(String.format("%-20s | %12s | %8d | %12d | %10d | %7.1f%%\n",
                "ИТОГО:",
                formatDuration(totalTime),
                totalFiles,
                totalProcessed,
                totalReports,
                totalFiles > 0 ? (totalProcessed * 100.0 / totalFiles) : 0));

        return sb.toString();
    }

    /**
     * Получить итоговую сводку программы
     */
    public static String getFinalSummary() {
        if (programStartTime == null) {
            return "Программа не была запущена";
        }

        Duration totalProgramTime = Duration.between(programStartTime, LocalDateTime.now());

        StringBuilder sb = new StringBuilder();
        sb.append("\n⭐").append("=".repeat(70)).append("⭐\n");
        sb.append("                      ИТОГОВАЯ СТАТИСТИКА ПРОГРАММЫ\n");
        sb.append("⭐").append("=".repeat(70)).append("⭐\n\n");

        sb.append(String.format("🏫 Обработано школ: %d\n", schoolMetrics.size()));
        sb.append(String.format("⏱️ Общее время выполнения: %s\n", formatDuration(totalProgramTime)));
        sb.append(String.format("🕐 Время запуска: %s\n",
                programStartTime.format(DateTimeFormatter.ofPattern("HH:mm:ss"))));
        sb.append(String.format("🕐 Время завершения: %s\n",
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"))));

        if (!schoolMetrics.isEmpty()) {
            // Считаем общую статистику
            int totalFiles = 0;
            int totalProcessed = 0;
            int totalReports = 0;
            Duration totalTime = Duration.ZERO;

            for (SchoolProcessingMetrics metrics : schoolMetrics.values()) {
                totalFiles += metrics.getFilesFound();
                totalProcessed += metrics.getFilesProcessed();
                totalReports += metrics.getReportsGenerated();
                totalTime = totalTime.plus(metrics.getTotalTime());
            }

            // Анализ производительности
            Optional<SchoolProcessingMetrics> fastest = schoolMetrics.values().stream()
                    .min(Comparator.comparing(SchoolProcessingMetrics::getTotalTime));
            Optional<SchoolProcessingMetrics> slowest = schoolMetrics.values().stream()
                    .max(Comparator.comparing(SchoolProcessingMetrics::getTotalTime));

            if (fastest.isPresent() && slowest.isPresent()) {
                sb.append(String.format("\n⚡ Самая быстрая школа: %s (%s)\n",
                        fastest.get().getSchoolName(), fastest.get().getFormattedDuration()));
                sb.append(String.format("🐢 Самая медленная школа: %s (%s)\n",
                        slowest.get().getSchoolName(), slowest.get().getFormattedDuration()));

                if (fastest.get().getTotalTime().toMillis() > 0) {
                    double ratio = (double) slowest.get().getTotalTime().toMillis() /
                            fastest.get().getTotalTime().toMillis();
                    sb.append(String.format("📈 Разница в скорости: %.1f раз\n", ratio));
                }
            }

            // Среднее время по фазам
            Duration avgFileFinding = calculateAverageDuration(
                    schoolMetrics.values().stream()
                            .map(SchoolProcessingMetrics::getFileFindingTime)
                            .toList()
            );

            Duration avgFileProcessing = calculateAverageDuration(
                    schoolMetrics.values().stream()
                            .map(SchoolProcessingMetrics::getFileProcessingTime)
                            .toList()
            );

            Duration avgReportGeneration = calculateAverageDuration(
                    schoolMetrics.values().stream()
                            .map(SchoolProcessingMetrics::getReportGenerationTime)
                            .toList()
            );

            sb.append("\n📊 Среднее время по фазам:\n");
            sb.append(String.format("   🔍 Поиск файлов: %s\n", formatDuration(avgFileFinding)));
            sb.append(String.format("   ⚙️ Обработка файлов: %s\n", formatDuration(avgFileProcessing)));
            sb.append(String.format("   📄 Генерация отчетов: %s\n", formatDuration(avgReportGeneration)));

            // Рекомендации по производительности
            sb.append("\n💡 РЕКОМЕНДАЦИИ ПО ОПТИМИЗАЦИИ:\n");

            if (slowest.isPresent() && slowest.get().getTotalTime().toSeconds() > 30) {
                sb.append("   • Добавьте индексы в БД для ускорения запросов\n");
            }

            if (avgFileFinding != null && avgFileFinding.toSeconds() > 5) {
                sb.append("   • Оптимизируйте поиск файлов (меньше вложенных папок)\n");
            }

            if (avgReportGeneration != null && avgReportGeneration.toSeconds() > 10) {
                sb.append("   • Упростите генерацию отчетов (меньше графиков/расчетов)\n");
            }

            // Исправлено: приведение типов
            long totalFailed = (long) totalFiles - totalProcessed;
            if (totalFiles > 0 && totalFailed > totalFiles * 0.1) { // больше 10% ошибок
                sb.append(String.format("   • Проверьте качество файлов (%d файлов не обработано)\n", totalFailed));
            }

            // Добавим процент успешной обработки
            if (totalFiles > 0) {
                double successRate = (totalProcessed * 100.0) / totalFiles;
                sb.append(String.format("   • Общая эффективность: %.1f%%\n", successRate));
                if (successRate < 90) {
                    sb.append("   • Рекомендуется проверить логи ошибок\n");
                }
            }
        }

        sb.append("\n✅ ПРОГРАММА ВЫПОЛНЕНА УСПЕШНО!\n");

        return sb.toString();
    }

    /**
     * Рассчитать среднюю длительность
     */
    private static Duration calculateAverageDuration(List<Duration> durations) {
        List<Duration> validDurations = durations.stream()
                .filter(Objects::nonNull)
                .toList();

        if (validDurations.isEmpty()) {
            return null;
        }

        Duration sum = Duration.ZERO;
        for (Duration duration : validDurations) {
            sum = sum.plus(duration);
        }

        return sum.dividedBy(validDurations.size());
    }

    /**
     * Форматировать длительность для вывода
     */
    private static String formatDuration(Duration duration) {
        if (duration == null || duration.isZero()) {
            return "        -";
        }
        if (duration.toSeconds() < 1) {
            return String.format("%4d мс", duration.toMillis());
        }
        long seconds = duration.toSeconds();
        long millis = duration.toMillisPart();
        return String.format("%2d.%03d сек", seconds, millis);
    }

    /**
     * Сохранить статистику в файл
     */
    public static void saveStatisticsToFile() {
        try {
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm"));
            java.nio.file.Path statsFile = Paths.get(STATISTIK_REPORT_FOLDER + timestamp + ".txt");

            List<String> lines = new ArrayList<>();
            lines.add("=".repeat(80));
            lines.add("СТАТИСТИКА ОБРАБОТКИ ОТЧЕТОВ ВСОКО");
            lines.add("Время генерации: " + LocalDateTime.now());
            lines.add("=".repeat(80));
            lines.add("");

            lines.add(getSchoolsStatistics());
            lines.add(getFinalSummary());

            // Детальная информация по каждой школе
            lines.add("\nДЕТАЛЬНАЯ ИНФОРМАЦИЯ ПО ШКОЛАМ:");
            lines.add("=".repeat(80));
            for (SchoolProcessingMetrics metrics : schoolMetrics.values()) {
                lines.add(String.format("\nШКОЛА: %s", metrics.getSchoolName()));
                lines.add(String.format("  Время начала: %s", metrics.getStartTime()));
                lines.add(String.format("  Время окончания: %s", metrics.getEndTime()));
                lines.add(String.format("  Длительность: %s", metrics.getFormattedDuration()));
                lines.add(String.format("  Найдено файлов: %d", metrics.getFilesFound()));
                lines.add(String.format("  Успешно обработано: %d", metrics.getFilesProcessed()));
                lines.add(String.format("  Сгенерировано отчетов: %d", metrics.getReportsGenerated()));
                lines.add(String.format("  Эффективность: %.1f%%", metrics.getSuccessRate()));

                if (metrics.getFileFindingTime() != null) {
                    lines.add(String.format("  Поиск файлов: %s", formatDuration(metrics.getFileFindingTime())));
                }
                if (metrics.getFileProcessingTime() != null) {
                    lines.add(String.format("  Обработка файлов: %s", formatDuration(metrics.getFileProcessingTime())));
                }
                if (metrics.getReportGenerationTime() != null) {
                    lines.add(String.format("  Генерация отчетов: %s", formatDuration(metrics.getReportGenerationTime())));
                }
            }

            java.nio.file.Files.write(statsFile, lines, java.nio.charset.StandardCharsets.UTF_8);
            log.info("📄 Статистика сохранена в файл: {}", statsFile.toAbsolutePath());

        } catch (Exception e) {
            log.error("⚠️ Не удалось сохранить статистику в файл", e);
        }
    }

    /**
     * Очистить статистику (для тестов)
     */
    public static void clear() {
        schoolMetrics.clear();
        programStartTime = null;
    }

    /**
     * Получить метрики для конкретной школы
     */
    public static SchoolProcessingMetrics getSchoolMetrics(String schoolName) {
        return schoolMetrics.get(schoolName);
    }

    /**
     * Получить все метрики
     */
    public static Map<String, SchoolProcessingMetrics> getAllMetrics() {
        return new HashMap<>(schoolMetrics);
    }

    /**
     * Проверить, есть ли статистика
     */
    public static boolean hasStatistics() {
        return !schoolMetrics.isEmpty();
    }
}