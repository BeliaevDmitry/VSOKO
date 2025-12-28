package org.school.analysis.service;

import org.school.analysis.model.ReportFile;

import java.io.File;
import java.util.List;

/**
 * Сервис поиска отчетов в папке
 */
public interface ReportFileFinderService {

    /**
     * Найти все Excel файлы в указанной папке
     * и создать объекты ReportFile
     */
    List<ReportFile> findReportFiles(String folderPath);

    /**
     * Найти файлы по паттерну имени
     */
    List<ReportFile> findFilesByPattern(String folderPath, String pattern);

    /**
     * Извлечь предмет и класс из имени файла
     */
    ReportFile extractMetadataFromFileName(File file);
}