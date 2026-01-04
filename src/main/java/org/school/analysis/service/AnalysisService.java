package org.school.analysis.service;

import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.model.dto.TestResultsDto;

import java.time.LocalDate;
import java.util.List;

/**
 * Сервис для анализа результатов тестов
 */
public interface AnalysisService {

    /**
     * Получить список всех тестов сгруппированных по предмету, классу, дате и учителю
     * @return список сводок по тестам
     */
    List<TestSummaryDto> getAllTestsSummary();

    /**
     * Получить список тестов с фильтрацией
     * @param subject предмет (опционально)
     * @param className класс (опционально)
     * @param startDate начальная дата (опционально)
     * @param endDate конечная дата (опционально)
     * @param teacher учитель (опционально)
     * @return отфильтрованный список тестов
     */
    List<TestSummaryDto> getTestsSummaryWithFilters(String subject, String className,
                                                    LocalDate startDate, LocalDate endDate,
                                                    String teacher);

    /**
     * Получить подробные результаты теста по указанным критериям
     * @param subject предмет
     * @param className класс
     * @param testDate дата теста
     * @param teacher учитель (опционально, для уточнения если есть дубли)
     * @return результаты теста со списком студентов
     */
    TestResultsDto getTestResults(String subject, String className, LocalDate testDate, String teacher);

    /**
     * Получить результаты теста по ID файла отчета
     * @param reportFileId ID файла отчета
     * @return результаты теста
     */
    TestResultsDto getTestResultsByReportFileId(String reportFileId);

    /**
     * Получить статистику по предмету
     * @param subject предмет
     * @return список тестов по предмету
     */
    List<TestSummaryDto> getTestsBySubject(String subject);

    /**
     * Получить статистику по классу
     * @param className класс
     * @return список тестов в классе
     */
    List<TestSummaryDto> getTestsByClass(String className);
}