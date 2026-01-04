package org.school.analysis.service;

import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.model.dto.TestSummaryDto;

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
     * Получить список результатов учеников по работе для анализа
     */
    List<StudentResultEntity> getResultTest(String school, String testType, String subject,
                                            String className);
}