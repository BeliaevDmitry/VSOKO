package org.school.analysis.service;

import org.school.analysis.model.ReportFile;

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
     * Создать структуру папок для предмета
     */
    String createSubjectFolderStructure(String subject) throws IOException;
}