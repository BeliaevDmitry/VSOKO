package org.school.analysis.service;

import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ReportFile;

import java.util.List;

/**
 * Сервис парсинга отчетов
 */
public interface ReportParserService {

    /**
     * Парсинг одного файла
     */
    ParseResult parseFile(ReportFile reportFile);

    /**
     * Парсинг списка файлов
     */
    List<ParseResult> parseFiles(List<ReportFile> reportFiles);

}