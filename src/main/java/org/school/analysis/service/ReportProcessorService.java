package org.school.analysis.service;

import lombok.Data;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ProcessingSummary;
import org.school.analysis.model.ReportFile;

import java.util.List;

/**
 * Главный сервис обработки отчетов
 * Координирует работу всех остальных сервисов
 */
public interface ReportProcessorService {

    /**
     * Полный цикл обработки
     */
    ProcessingSummary processAll(String folderPath);

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
}