package org.school.analysis.service;

import lombok.Data;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;

import java.util.List;

/**
 * Главный сервис обработки отчетов
 * Координирует работу всех остальных сервисов
 */
public interface ReportProcessorService {

    /**
     * 1. Найти файлы в папке
     */
    List<ReportFile> findReports(String folderPath);

    /**
     * 2. Парсинг найденных файлов
     */
    List<ParseResult> parseReports(List<ReportFile> reportFiles);

    /**
     * 3. Сохранение результатов в БД
     */
    List<ReportFile> saveResultsToDatabase(List<ParseResult> parseResults);

    /**
     * 4. Перемещение обработанных файлов
     */
    List<ReportFile> moveProcessedFiles(List<ReportFile> successfullyProcessedFiles);

    /**
     * Полный цикл обработки
     */
    ProcessingSummary processAll(String folderPath);

    @Data
    class ProcessingSummary {
        private int totalFilesFound;
        private int successfullyParsed;
        private int successfullySaved;
        private int successfullyMoved;
        private int totalStudentsProcessed;
        private List<ReportFile> failedFiles;
    }
}