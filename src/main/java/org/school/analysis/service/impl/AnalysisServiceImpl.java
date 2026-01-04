package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.repository.ReportFileRepository;
import org.school.analysis.repository.StudentResultRepository;
import org.school.analysis.service.AnalysisService;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.domain.Specification;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import jakarta.persistence.criteria.Join;
import jakarta.persistence.criteria.JoinType;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
@Transactional(readOnly = true)
public class AnalysisServiceImpl implements AnalysisService {

    private final ReportFileRepository reportFileRepository;
    private final StudentResultRepository studentResultRepository;

    @Override
    public List<TestSummaryDto> getAllTestsSummary() {
        log.info("Получение сводки по всем тестам");

        // Сортируем по дате (сначала новые), затем по предмету и классу
        List<ReportFileEntity> reportFiles = reportFileRepository.findAll(
                Sort.by(Sort.Direction.DESC, "testDate")
                        .and(Sort.by("subject"))
                        .and(Sort.by("className"))
        );

        // Конвертируем в DTO и рассчитываем средний балл для каждого теста
        return reportFiles.stream()
                .map(this::convertToTestSummaryDto)
                .filter(Objects::nonNull)
                .collect(Collectors.toList());
    }

    public List<StudentResultEntity> getResultTest(String school, String testType, String subject,
                                                   String className) {
        //получаем список результатов учеников
        Map<Integer, Integer> taskScores;
        return null;
    }

    /**
     * Конвертирует ReportFileEntity в TestSummaryDto с расчетом статистики
     */
    /**
     * Конвертирует ReportFileEntity в TestSummaryDto с расчетом статистики
     */
    private TestSummaryDto convertToTestSummaryDto(ReportFileEntity reportFile) {
        if (reportFile == null) {
            log.warn("ReportFileEntity is null");
            return null;
        }

        try {
            // Получаем всех студентов для этого отчета
            List<StudentResultEntity> allStudents = getStudentsForReport(reportFile);

            // Считаем присутствующих и отсутствующих
            long presentCount = allStudents.stream()
                    .filter(s -> s.getPresence() != null && "Был".equalsIgnoreCase(s.getPresence()))
                    .count();

            long absentCount = allStudents.stream()
                    .filter(s -> s.getPresence() != null && ("Не был".equalsIgnoreCase(s.getPresence()) ||
                            "Отсутствовал".equalsIgnoreCase(s.getPresence()) ||
                            "Отсутствовала".equalsIgnoreCase(s.getPresence())))
                    .count();

            // Получаем средний балл для этого теста (только по присутствующим)
            Double averageScore = calculateAverageScoreForReport(reportFile, true);

            // Рассчитываем максимальный общий балл из JSON
            Integer maxTotalScore = calculateMaxTotalScoreFromJson(reportFile.getMaxScoresJson());

            // Вычисляем количество учеников, которые писали тест
            int studentsCount = (int) presentCount; // Только присутствовавшие писали

            // Получаем общее количество учеников в классе из reportFile
            int classSize = reportFile.getStudentCount() != null ? reportFile.getStudentCount() : 0;

            return TestSummaryDto.builder()
                    .school(reportFile.getSchool())
                    .subject(reportFile.getSubject())
                    .className(reportFile.getClassName())
                    .testDate(reportFile.getTestDate())
                    .testType(reportFile.getTestType())
                    .teacher(reportFile.getTeacher())
                    .studentsCount(studentsCount)                 // Количество писавших (присутствовавших)
                    .classSize(classSize)                         // Всего учеников в классе
                    .taskCount(reportFile.getTaskCount())
                    .maxTotalScore(maxTotalScore)                 // Вычисляем из JSON
                    .averageScore(averageScore)                   // Средний балл (по присутствующим)
                    .fileName(reportFile.getFileName())
                    // Дополнительные поля для совместимости
                    .studentsTotal(classSize)                     // Всего учеников в классе
                    .studentsPresent(studentsCount)               // Присутствовало на тесте
                    .studentsAbsent((int) absentCount)            // Отсутствовало на тесте
                    .build();
        } catch (Exception e) {
            log.error("Ошибка при конвертации ReportFileEntity в TestSummaryDto: {}", e.getMessage(), e);
            return null;
        }
    }

    /**
     * Рассчитывает максимальный общий балл из JSON
     */
    private Integer calculateMaxTotalScoreFromJson(String maxScoresJson) {
        if (maxScoresJson == null || maxScoresJson.trim().isEmpty()) {
            log.warn("maxScoresJson is null or empty");
            return 0;
        }
        return JsonScoreUtils.calculateTotalScore(maxScoresJson);
    }

    /**
     * Рассчитывает средний балл для отчета
     *
     * @param onlyPresent true - считать только присутствующих, false - считать всех
     */
    private Double calculateAverageScoreForReport(ReportFileEntity reportFile, boolean onlyPresent) {
        try {
            List<StudentResultEntity> students = studentResultRepository.findAll(
                    (root, query, cb) -> {
                        Join<StudentResultEntity, ReportFileEntity> reportFileJoin =
                                root.join("reportFile", JoinType.INNER);
                        return cb.equal(reportFileJoin.get("id"), reportFile.getId());
                    }
            );

            if (students.isEmpty()) {
                log.debug("Нет студентов для отчета {}", reportFile.getFileName());
                return 0.0;
            }

            // Фильтруем если нужно только присутствующих
            List<StudentResultEntity> studentsToCalculate = onlyPresent ?
                    students.stream()
                            .filter(s -> s.getPresence() != null && "Был".equalsIgnoreCase(s.getPresence()))
                            .collect(Collectors.toList()) :
                    students;

            // Считаем средний балл
            double totalScore = studentsToCalculate.stream()
                    .map(StudentResultEntity::getTotalScore)
                    .filter(Objects::nonNull)
                    .mapToInt(Integer::intValue)
                    .sum();

            long count = studentsToCalculate.stream()
                    .filter(s -> s.getTotalScore() != null)
                    .count();

            return count > 0 ? Math.round((totalScore / count) * 100.0) / 100.0 : 0.0;

        } catch (Exception e) {
            log.error("Ошибка расчета среднего балла для отчета {}: {}",
                    reportFile.getFileName(), e.getMessage());
            return 0.0;
        }
    }

    /**
     * Получает всех студентов для отчета (вспомогательный метод)
     */
    private List<StudentResultEntity> getStudentsForReport(ReportFileEntity reportFile) {
        try {
            Specification<StudentResultEntity> studentSpec = (root, query, criteriaBuilder) -> {
                Join<StudentResultEntity, ReportFileEntity> reportFileJoin =
                        root.join("reportFile", JoinType.INNER);
                return criteriaBuilder.equal(reportFileJoin.get("id"), reportFile.getId());
            };

            return studentResultRepository.findAll(studentSpec);
        } catch (Exception e) {
            log.error("Ошибка получения студентов для отчета {}: {}",
                    reportFile.getFileName(), e.getMessage());
            return List.of();
        }
    }


    /**
     * Получение детальной статистики по конкретному тесту
     */
    public TestSummaryDto getTestSummaryById(String reportFileId) {
        log.info("Получение сводки по тесту с ID: {}", reportFileId);

        try {
            ReportFileEntity reportFile = reportFileRepository.findById(java.util.UUID.fromString(reportFileId))
                    .orElseThrow(() -> new IllegalArgumentException("Тест с ID " + reportFileId + " не найден"));

            return convertToTestSummaryDto(reportFile);
        } catch (Exception e) {
            log.error("Ошибка при получении теста по ID {}: {}", reportFileId, e.getMessage());
            throw new RuntimeException("Не удалось получить информацию о тесте", e);
        }
    }

    /**
     * Получение распределения баллов для визуализации
     */
    public Map<Integer, Integer> getScoreDistribution(String reportFileId) {
        log.info("Получение распределения баллов для теста с ID: {}", reportFileId);

        try {
            ReportFileEntity reportFile = reportFileRepository.findById(java.util.UUID.fromString(reportFileId))
                    .orElseThrow(() -> new IllegalArgumentException("Тест с ID " + reportFileId + " не найден"));

            List<StudentResultEntity> students = getStudentsForReport(reportFile);

            // Группируем студентов по баллам
            return students.stream()
                    .filter(s -> s.getTotalScore() != null && "Был".equalsIgnoreCase(s.getPresence()))
                    .collect(Collectors.groupingBy(
                            StudentResultEntity::getTotalScore,
                            Collectors.collectingAndThen(Collectors.counting(), Long::intValue)
                    ));
        } catch (Exception e) {
            log.error("Ошибка при получении распределения баллов для теста {}: {}",
                    reportFileId, e.getMessage());
            return Map.of();
        }
    }
}