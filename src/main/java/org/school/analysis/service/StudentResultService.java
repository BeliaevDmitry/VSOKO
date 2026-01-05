package org.school.analysis.service;

import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;

import java.util.List;
import java.util.UUID;

public interface StudentResultService {

    /**
     * Сохраняет все результаты студентов из файла
     */
    int saveAll(ReportFile reportFile, List<StudentResult> studentResults);

    /**
     * Получает файл отчета по ID
     */
    ReportFile getReportFileById(UUID id);

    /**
     * Получает результаты студентов по ID файла отчета
     */
    List<StudentResult> getStudentResultsByReportFileId(UUID reportFileId);

    /**
     * Получает количество студентов для файла отчета
     */
    long countByReportFileId(UUID reportFileId);
}