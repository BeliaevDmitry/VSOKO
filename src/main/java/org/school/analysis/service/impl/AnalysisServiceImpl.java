package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.model.dto.TestResultsDto;
import org.school.analysis.repository.ReportFileRepository;
import org.school.analysis.repository.StudentResultRepository;
import org.school.analysis.service.AnalysisService;
import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.domain.Specification;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import jakarta.persistence.criteria.Join;
import jakarta.persistence.criteria.JoinType;
import jakarta.persistence.criteria.Predicate;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
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
                .collect(Collectors.toList());
    }

    @Override
    public List<TestSummaryDto> getTestsSummaryWithFilters(String subject, String className,
                                                           LocalDate startDate, LocalDate endDate,
                                                           String teacher, String school) {
        log.info("Получение тестов с фильтрами");

        Specification<ReportFileEntity> spec = buildReportFileSpecification(
                subject, className, startDate, endDate, teacher, school);

        List<ReportFileEntity> reportFiles = reportFileRepository.findAll(
                spec,
                Sort.by(Sort.Direction.DESC, "testDate")
                        .and(Sort.by("subject"))
                        .and(Sort.by("className"))
        );

        return reportFiles.stream()
                .map(this::convertToTestSummaryDto)
                .collect(Collectors.toList());
    }

    @Override
    public TestResultsDto getTestResults(String subject, String className,
                                         LocalDate testDate, String teacher) {
        // TODO: Реализовать позже
        return null;
    }

    @Override
    public Double getAverageScoreForTest(String subject, String className,
                                         LocalDate testDate, String teacher) {
        // TODO: Реализовать позже
        return null;
    }

    /**
     * Конвертирует ReportFileEntity в TestSummaryDto с расчетом статистики
     */
    private TestSummaryDto convertToTestSummaryDto(ReportFileEntity reportFile) {
        // Получаем всех студентов для этого отчета
        List<StudentResultEntity> allStudents = getStudentsForReport(reportFile);

        // Считаем присутствующих и отсутствующих
        long presentCount = allStudents.stream()
                .filter(s -> s.getPresence() != null && "Был".equalsIgnoreCase(s.getPresence()))
                .count();

        long absentCount = allStudents.stream()
                .filter(s -> s.getPresence() != null && "Не был".equalsIgnoreCase(s.getPresence()))
                .count();

        // Получаем средний балл для этого теста (только по присутствующим)
        Double averageScore = calculateAverageScoreForReport(reportFile, true);

        // Получаем количество учеников в классе (из справочника или заглушка)
        Integer classSize = getClassSize(reportFile.getClassName());

        // Если classSize не получен из справочника, используем данные из теста
        if (classSize == null || classSize == 0) {
            classSize = (int) (presentCount + absentCount);
        }

        return TestSummaryDto.builder()
                .school(reportFile.getSchool())
                .subject(reportFile.getSubject())
                .className(reportFile.getClassName())
                .testDate(reportFile.getTestDate())
                .testType(reportFile.getTestType())
                .teacher(reportFile.getTeacher())
                .studentsTotal((int) (presentCount + absentCount)) // Всего в файле
                .studentsPresent((int) presentCount)               // Присутствовало
                .studentsAbsent((int) absentCount)                 // Отсутствовало
                .classSize(reportFile.getStudentCount())           // Всего в классе (из справочника)
                .taskCount(reportFile.getTaskCount())
                .maxTotalScore(reportFile.getMaxTotalScore())
                .averageScore(averageScore)                       // Средний балл (по присутствующим)
                .fileName(reportFile.getFileName())
                .build();
    }

    /**
     * Рассчитывает средний балл для отчета
     * @param onlyPresent true - считать только присутствующих, false - считать всех
     */
    private Double calculateAverageScoreForReport(ReportFileEntity reportFile, boolean onlyPresent) {
        try {
            // Находим всех студентов для этого отчета
            Specification<StudentResultEntity> studentSpec = (root, query, criteriaBuilder) -> {
                Join<StudentResultEntity, ReportFileEntity> reportFileJoin =
                        root.join("reportFile", JoinType.INNER);
                return criteriaBuilder.equal(reportFileJoin.get("id"), reportFile.getId());
            };

            List<StudentResultEntity> students = studentResultRepository.findAll(studentSpec);

            if (students.isEmpty()) {
                return 0.0;
            }

            // Фильтруем студентов если нужно считать только присутствующих
            List<StudentResultEntity> studentsToCalculate = students;
            if (onlyPresent) {
                studentsToCalculate = students.stream()
                        .filter(s -> s.getPresence() != null && "Был".equalsIgnoreCase(s.getPresence()))
                        .collect(Collectors.toList());
            }

            // Рассчитываем средний балл
            double totalScore = studentsToCalculate.stream()
                    .filter(s -> s.getTotalScore() != null)
                    .mapToInt(StudentResultEntity::getTotalScore)
                    .sum();

            long count = studentsToCalculate.stream()
                    .filter(s -> s.getTotalScore() != null)
                    .count();

            return count > 0 ? Math.round((totalScore / count) * 100.0) / 100.0 : 0.0;

        } catch (Exception e) {
            log.error("Ошибка расчета среднего балла для отчета {}: {}",
                    reportFile.getId(), e.getMessage());
            return 0.0;
        }
    }

    /**
     * Получает всех студентов для отчета (вспомогательный метод)
     */
    private List<StudentResultEntity> getStudentsForReport(ReportFileEntity reportFile) {
        Specification<StudentResultEntity> studentSpec = (root, query, criteriaBuilder) -> {
            Join<StudentResultEntity, ReportFileEntity> reportFileJoin =
                    root.join("reportFile", JoinType.INNER);
            return criteriaBuilder.equal(reportFileJoin.get("id"), reportFile.getId());
        };

        return studentResultRepository.findAll(studentSpec);
    }

    /**
     * Заглушка для получения количества учеников в классе
     * TODO: Реализовать получение из таблицы классов/учеников
     */


    /**
     * Рассчитывает средний балл для отчета
     */
    private Double calculateAverageScoreForReport(ReportFileEntity reportFile) {
        try {
            // Находим всех студентов для этого отчета
            Specification<StudentResultEntity> studentSpec = (root, query, criteriaBuilder) -> {
                Join<StudentResultEntity, ReportFileEntity> reportFileJoin =
                        root.join("reportFile", JoinType.INNER);
                return criteriaBuilder.equal(reportFileJoin.get("id"), reportFile.getId());
            };

            List<StudentResultEntity> students = studentResultRepository.findAll(studentSpec);

            if (students.isEmpty()) {
                return 0.0;
            }

            // Рассчитываем средний балл
            double totalScore = students.stream()
                    .filter(s -> s.getTotalScore() != null)
                    .mapToInt(StudentResultEntity::getTotalScore)
                    .sum();

            long count = students.stream()
                    .filter(s -> s.getTotalScore() != null)
                    .count();

            return count > 0 ? Math.round((totalScore / count) * 100.0) / 100.0 : 0.0;

        } catch (Exception e) {
            log.error("Ошибка расчета среднего балла для отчета {}: {}",
                    reportFile.getId(), e.getMessage());
            return 0.0;
        }
    }

    /**
     * Заглушка для получения количества учеников в классе
     * TODO: Реализовать получение из таблицы классов/учеников
     */
    private Integer getClassSize(String className) {
        // Временная логика - предполагаем стандартный размер класса
        // Можно добавить конфигурацию в AppConfig
        return 25; // Средний размер класса
    }

    /**
     * Строит спецификацию для фильтрации файлов отчетов
     */
    private Specification<ReportFileEntity> buildReportFileSpecification(
            String subject, String className,
            LocalDate startDate, LocalDate endDate,
            String teacher, String school) {

        return (root, query, criteriaBuilder) -> {
            List<Predicate> predicates = new ArrayList<>();

            if (subject != null && !subject.trim().isEmpty()) {
                predicates.add(criteriaBuilder.equal(root.get("subject"), subject));
            }

            if (className != null && !className.trim().isEmpty()) {
                predicates.add(criteriaBuilder.equal(root.get("className"), className));
            }

            if (startDate != null) {
                predicates.add(criteriaBuilder.greaterThanOrEqualTo(root.get("testDate"), startDate));
            }

            if (endDate != null) {
                predicates.add(criteriaBuilder.lessThanOrEqualTo(root.get("testDate"), endDate));
            }

            if (teacher != null && !teacher.trim().isEmpty()) {
                predicates.add(criteriaBuilder.like(
                        criteriaBuilder.lower(root.get("teacher")),
                        "%" + teacher.toLowerCase() + "%"
                ));
            }

            if (school != null && !school.trim().isEmpty()) {
                predicates.add(criteriaBuilder.equal(root.get("school"), school));
            }

            // Только обработанные файлы
            predicates.add(criteriaBuilder.equal(root.get("status"), "PROCESSED"));

            return criteriaBuilder.and(predicates.toArray(new Predicate[0]));
        };
    }
}