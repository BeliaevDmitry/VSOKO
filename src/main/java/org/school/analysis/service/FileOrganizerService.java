package org.school.analysis.service;

import org.school.analysis.model.ReportFile;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Сервис организации файлов по папкам
 */
public interface FileOrganizerService {

    /**
     * Переместить файл в папку предмета
     */
    boolean moveToSubjectFolder(ReportFile reportFile) throws IOException;

    /**
     * Переместить несколько файлов
     */
    List<ReportFile> moveFilesToSubjectFolders(List<ReportFile> reportFiles);

    /**
     * Найти все Excel файлы в указанной папке
     * и создать объекты ReportFile
     */
    List<ReportFile> findReportFiles(String folderPath);

    /**
     * Извлечь предмет и класс из имени файла
     */
    ReportFile extractMetadataFromFileName(File file);

}