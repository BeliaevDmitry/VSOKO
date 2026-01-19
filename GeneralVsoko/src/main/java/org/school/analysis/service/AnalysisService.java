package org.school.analysis.service;

import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TestSummaryDto;

import java.util.List;
import java.util.Map;

public interface AnalysisService {

    List<TestSummaryDto> getAllTestsSummary(String schoolName, String currentAcademicYear);

    /**
     * Получить детальные результаты студентов для теста
     */
    List<StudentDetailedResultDto> getStudentDetailedResults(String reportFileId);

    /**
     * Получить статистику по заданиям для теста
     */
    Map<Integer, TaskStatisticsDto> getTaskStatistics(String reportFileId);

    /**
     * Получить тесты по учителю
     */
    List<TestSummaryDto> getTestsByTeacher(String teacherName, String school, String currentAcademicYear);

    /**
     * Получить список уникальных учителей
     */
    List<String> getAllTeachers(String school, String currentAcademicYear);
}