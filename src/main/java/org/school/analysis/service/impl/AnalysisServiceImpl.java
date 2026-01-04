package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.model.dto.StudentTestResultDto;
import org.school.analysis.model.dto.TestResultsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.repository.ReportFileRepository;
import org.school.analysis.repository.StudentResultRepository;
import org.school.analysis.service.AnalysisService;
import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.domain.Specification;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.StringUtils;

import jakarta.persistence.criteria.Join;
import jakarta.persistence.criteria.JoinType;
import jakarta.persistence.criteria.Predicate;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.UUID;
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

        List<ReportFileEntity> reportFiles = reportFileRepository.findAll(
                Sort.by(Sort.Direction.DESC, "testDate")
                        .and(Sort.by("subject"))
                        .and(Sort.by("className"))
        );

        return convertToTestSummaryDto(reportFiles);
    }

    @Override
    public List<TestSummaryDto> getTestsSummaryWithFilters(String subject, String className,
                                                           LocalDate startDate, LocalDate endDate,
                                                           String teacher) {
        log.info("Получение тестов с фильтрами: subject={}, className={}, startDate={}, endDate={}, teacher={}",
                subject, className, startDate, endDate, teacher);

        Specification<ReportFileEntity> specification = buildReportFileSpecification(
                subject, className, startDate, endDate, teacher, null
        );

        List<ReportFileEntity> reportFiles = reportFileRepository.findAll(
                specification,
                Sort.by(Sort.Direction.DESC, "testDate")
                        .and(Sort.by("subject"))
                        .and(Sort.by("className"))
        );

        return convertToTestSummaryDto(reportFiles);
    }

    @Override
    public TestResultsDto getTestResults(String subject, String className, LocalDate testDate, String teacher) {
        log.info("Получение результатов теста: {} {} {} {}", subject, className, testDate, teacher);

        Specification<ReportFileEntity> specification = buildReportFileSpecification(
                subject, className, null, null, teacher, testDate
        );

        List<ReportFileEntity> reportFiles = reportFileRepository.findAll(specification);

        if (reportFiles.isEmpty()) {
            log.warn("Тест не найден: {} {} {} {}", subject, className, testDate, teacher);
            return TestResultsDto.builder().build();
        }

        // Если найдено несколько, берем первый
        ReportFileEntity reportFile = reportFiles.get(0);
        return getTestResultsByReportFile(reportFile);
    }

    @Override
    public TestResultsDto getTestResultsByReportFileId(String reportFileId) {
        log.info("Получение результатов теста по ID файла: {}", reportFileId);

        try {
            UUID id = UUID.fromString(reportFileId);
            ReportFileEntity reportFile = reportFileRepository.findById(id)
                    .orElseThrow(() -> new IllegalArgumentException("Файл отчета не найден: " + reportFileId));

            return getTestResultsByReportFile(reportFile);

        } catch (IllegalArgumentException e) {
            log.error("Неверный формат UUID: {}", reportFileId, e);
            return TestResultsDto.builder().build();
        }
    }

    @Override
    public List<TestSummaryDto> getTestsBySubject(String subject) {
        log.info("Получение тестов по предмету: {}", subject);
        return getTestsSummaryWithFilters(subject, null, null, null, null);
    }

    @Override
    public List<TestSummaryDto> getTestsByClass(String className) {
        log.info("Получение тестов по классу: {}", className);
        return getTestsSummaryWithFilters(null, className, null, null, null);
    }

    /**
     * Строит спецификацию для фильтрации файлов отчетов
     */
    private Specification<ReportFileEntity> buildReportFileSpecification(
            String subject, String className,
            LocalDate startDate, LocalDate endDate,
            String teacher, LocalDate exactDate) {

        return (root, query, criteriaBuilder) -> {
            List<Predicate> predicates = new ArrayList<>();

            // Фильтр по предмету
            if (StringUtils.hasText(subject)) {
                predicates.add(criteriaBuilder.equal(root.get("subject"), subject));
            }

            // Фильтр по классу
            if (StringUtils.hasText(className)) {
                predicates.add(criteriaBuilder.equal(root.get("className"), className));
            }

            // Фильтр по дате (точная дата ИЛИ диапазон)
            if (exactDate != null) {
                predicates.add(criteriaBuilder.equal(root.get("testDate"), exactDate));
            } else {
                if (startDate != null) {
                    predicates.add(criteriaBuilder.greaterThanOrEqualTo(root.get("testDate"), startDate));
                }
                if (endDate != null) {
                    predicates.add(criteriaBuilder.lessThanOrEqualTo(root.get("testDate"), endDate));
                }
            }

            // Фильтр по учителю
            if (StringUtils.hasText(teacher)) {
                predicates.add(criteriaBuilder.like(
                        criteriaBuilder.lower(root.get("teacher")),
                        "%" + teacher.toLowerCase() + "%"
                ));
            }

            // Только обработанные файлы
            predicates.add(criteriaBuilder.equal(root.get("status"), "PROCESSED"));

            return criteriaBuilder.and(predicates.toArray(new Predicate[0]));
        };
    }

    /**
     * Конвертирует список сущностей ReportFileEntity в DTO
     */
    private List<TestSummaryDto> convertToTestSummaryDto(List<ReportFileEntity> reportFiles) {
        return reportFiles.stream()
                .map(this::convertToTestSummaryDto)
                .collect(Collectors.toList());
    }

    /**
     * Конвертирует одну сущность ReportFileEntity в DTO
     */
    private TestSummaryDto convertToTestSummaryDto(ReportFileEntity reportFile) {
        return TestSummaryDto.builder()
                .subject(reportFile.getSubject())
                .className(reportFile.getClassName())
                .testDate(reportFile.getTestDate())
                .teacher(reportFile.getTeacher())
                .school(reportFile.getSchool())
                .studentCount(reportFile.getStudentCount())
                .taskCount(reportFile.getTaskCount())
                .maxTotalScore(reportFile.getMaxTotalScore())
                .testType(reportFile.getTestType())
                .fileName(reportFile.getFileName())
                .build();
    }

    /**
     * Получает результаты теста по сущности ReportFileEntity
     */
    private TestResultsDto getTestResultsByReportFile(ReportFileEntity reportFile) {
        // Создаем спецификацию для студентов этого отчета
        Specification<StudentResultEntity> studentSpec = (root, query, criteriaBuilder) -> {
            Join<StudentResultEntity, ReportFileEntity> reportFileJoin = root.join("reportFile", JoinType.INNER);
            return criteriaBuilder.equal(reportFileJoin.get("id"), reportFile.getId());
        };

        // Получаем всех студентов для этого отчета
        List<StudentResultEntity> studentResults = studentResultRepository.findAll(
                studentSpec,
                Sort.by("fio")
        );

        // Конвертируем студентов в DTO
        List<StudentTestResultDto> studentDtos = studentResults.stream()
                .map(this::convertToStudentTestResultDto)
                .collect(Collectors.toList());

        // Сортируем студентов по баллам (от большего к меньшему) и добавляем позиции
        studentDtos.sort(Comparator.comparing(StudentTestResultDto::getTotalScore,
                Comparator.nullsLast(Comparator.reverseOrder())));

        for (int i = 0; i < studentDtos.size(); i++) {
            studentDtos.get(i).setPositionInClass(i + 1);
        }

        // Создаем итоговый DTO
        return TestResultsDto.builder()
                .testSummary(convertToTestSummaryDto(reportFile))
                .studentResults(studentDtos)
                .build();
    }

    /**
     * Конвертирует сущность StudentResultEntity в DTO
     */
    private StudentTestResultDto convertToStudentTestResultDto(StudentResultEntity student) {
        return StudentTestResultDto.builder()
                .fio(student.getFio())
                .presence(student.getPresence())
                .variant(student.getVariant())
                .totalScore(student.getTotalScore())
                .percentageScore(student.getPercentageScore())
                .build();
    }
}