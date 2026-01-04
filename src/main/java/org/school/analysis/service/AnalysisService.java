package org.school.analysis.service;

import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.model.dto.TestResultsDto;

import java.time.LocalDate;
import java.util.List;

public interface AnalysisService {

    /**
     * Получить список всех тестов сгруппированных по предмету, классу, дате, учителю
     * Формат: Школа, Предмет, Класс, Дата теста, Тип теста, Учитель,
     * Кол-во учеников писавших, Кол-во учеников в классе, Кол-во заданий теста,
     * Макс. балл, Средний балл теста
     */
    List<TestSummaryDto> getAllTestsSummary();

    /**
     * Получить список тестов с фильтрами
     */
    List<TestSummaryDto> getTestsSummaryWithFilters(String subject, String className,
                                                    LocalDate startDate, LocalDate endDate,
                                                    String teacher, String school);

    /**
     * Получить подробные результаты теста
     */
    TestResultsDto getTestResults(String subject, String className,
                                  LocalDate testDate, String teacher);

    /**
     * Получить средний балл по тесту
     */
    Double getAverageScoreForTest(String subject, String className,
                                  LocalDate testDate, String teacher);
}