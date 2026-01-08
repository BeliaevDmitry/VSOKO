package org.school.analysis.model;

import lombok.Data;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

// класс статистики
@Data
public class ProcessingSummary {
    private int totalFilesFound;
    private int successfullyParsed;
    private int successfullySaved;
    private int successfullyMoved;
    private int generatedReportsCount;
    private List<File> generatedReportFiles = new ArrayList<>();
    private String reportGenerationError;

    // Добавьте это поле
    private List<ReportFile> failedFiles = new ArrayList<>();

    public int getFailedToParse() {
        return totalFilesFound - successfullyParsed;
    }

    public int getFailedToSave() {
        return successfullyParsed - successfullySaved;
    }

    public int getFailedToMove() {
        return successfullySaved - successfullyMoved;
    }

    @Override
    public String toString() {
        return String.format(
                "Обработка завершена: Найдено=%d, Распарсено=%d, Сохранено=%d, Перемещено=%d, Отчетов=%d, Неудачных=%d",
                totalFilesFound, successfullyParsed, successfullySaved, successfullyMoved,
                generatedReportsCount, failedFiles.size()
        );
    }

    // Дополнительные методы для удобства

    /**
     * Получить количество неудачных файлов
     */
    public int getFailedFilesCount() {
        return failedFiles.size();
    }

    /**
     * Добавить неудачный файл
     */
    public void addFailedFile(ReportFile failedFile) {
        if (failedFile != null) {
            failedFiles.add(failedFile);
        }
    }

    /**
     * Добавить несколько неудачных файлов
     */
    public void addAllFailedFiles(List<ReportFile> failedFilesList) {
        if (failedFilesList != null) {
            failedFiles.addAll(failedFilesList);
        }
    }

    /**
     * Получить статистику по статусам ошибок
     */
    public String getFailureStatistics() {
        if (failedFiles.isEmpty()) {
            return "Все файлы обработаны успешно";
        }

        StringBuilder stats = new StringBuilder();
        stats.append("Неудачные файлы (").append(failedFiles.size()).append("):\n");

        // Группировка по статусам ошибок
        int parsingErrors = 0;
        int savingErrors = 0;
        int movingErrors = 0;
        int validationErrors = 0;

        for (ReportFile file : failedFiles) {
            if (file.getStatus() != null) {
                switch (file.getStatus()) {
                    case ERROR_PARSING:
                        parsingErrors++;
                        break;
                    case ERROR_SAVING:
                        savingErrors++;
                        break;
                    case ERROR_MOVING:
                        movingErrors++;
                        break;
                    case ERROR_VALIDATION:
                        validationErrors++;
                        break;
                    default:
                        break;
                }
            }
        }

        if (parsingErrors > 0) {
            stats.append("  • Ошибки парсинга: ").append(parsingErrors).append("\n");
        }
        if (savingErrors > 0) {
            stats.append("  • Ошибки сохранения: ").append(savingErrors).append("\n");
        }
        if (movingErrors > 0) {
            stats.append("  • Ошибки перемещения: ").append(movingErrors).append("\n");
        }
        if (validationErrors > 0) {
            stats.append("  • Ошибки валидации: ").append(validationErrors).append("\n");
        }

        return stats.toString();
    }
    // Добавить эти методы
    public void incrementSuccessfullyParsed(int count) {
        this.successfullyParsed += count;
    }

    public void incrementSuccessfullySaved(int count) {
        this.successfullySaved += count;
    }

    public void incrementSuccessfullyMoved(int count) {
        this.successfullyMoved += count;
    }

    public void incrementTotalFilesFound(int count) {
        this.totalFilesFound += count;
    }

    public void incrementGeneratedReportsCount(int count) {
        this.generatedReportsCount += count;
    }
}