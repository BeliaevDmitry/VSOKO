package org.school.analysis.mapper;

import org.school.analysis.model.entity.ReportFileEntity;
import org.school.analysis.model.entity.StudentResultEntity;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.JsonScoreUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.time.LocalDateTime;

@Component
public class ReportMapper {
    private static final Logger log = LoggerFactory.getLogger(ReportMapper.class);

    private static final int SUBJECT_MAX = 100;
    private static final int CLASS_NAME_MAX = 50;
    private static final int TEST_TYPE_MAX = 50;
    private static final int FILE_NAME_MAX = 255;
    private static final int FIO_MAX = 200;
    private static final int VARIANT_MAX = 100;
    private static final int PRESENCE_MAX = 50;

    public ReportFileEntity toEntity(ReportFile model) {
        if (model == null) {
            return null;
        }

        return ReportFileEntity.builder()
                .filePath(model.getFile() != null ? model.getFile().getAbsolutePath() : "")
                .fileName(truncate(model.getFileName(), FILE_NAME_MAX, "report_files.file_name"))
                .fileHash("") // Вычисляется отдельно
                .subject(truncate(model.getSubject(), SUBJECT_MAX, "report_files.subject"))
                .className(truncate(model.getClassName(), CLASS_NAME_MAX, "report_files.class_name"))
                .status(model.getStatus())
                .processedAt(model.getProcessedAt())
                .errorMessage(model.getErrorMessage())
                .studentCount(model.getStudentCount())
                .testDate(model.getTestDate())
                .teacher(truncate(model.getTeacher(), FIO_MAX, "report_files.teacher"))
                .schoolName(model.getSchoolName() != null ? model.getSchoolName() : "ГБОУ №7")
                .academicYear(model.getAcademicYear() != null ? model.getAcademicYear() : "2025-2026")
                .taskCount(model.getMaxScores() != null ? model.getMaxScores().size() : 0)
                .testType(truncate(model.getTestType(), TEST_TYPE_MAX, "report_files.test_type"))
                .comment(model.getComment())
                .maxScoresJson(JsonScoreUtils.mapToJson(model.getMaxScores()))
                .createdAt(LocalDateTime.now())
                .updatedAt(LocalDateTime.now())
                .build();
    }

    public ReportFile toModel(ReportFileEntity entity) {
        if (entity == null) {
            return null;
        }

        ReportFile model = new ReportFile();
        model.setSubject(entity.getSubject());
        model.setClassName(entity.getClassName());
        model.setStatus(entity.getStatus());
        model.setProcessedAt(entity.getProcessedAt());
        model.setErrorMessage(entity.getErrorMessage());
        model.setStudentCount(entity.getStudentCount());
        model.setTestDate(entity.getTestDate());
        model.setTeacher(entity.getTeacher());
        model.setSchoolName(entity.getSchoolName());
        model.setTaskCount(entity.getTaskCount());
        model.setMaxScores(JsonScoreUtils.jsonToMap(entity.getMaxScoresJson()));
        model.setTestType(entity.getTestType());
        model.setComment(entity.getComment());
        model.setAcademicYear(entity.getAcademicYear());
        return model;
    }

    public StudentResultEntity toEntity(StudentResult model, ReportFileEntity reportFile) {
        if (model == null) {
            return null;
        }

        return StudentResultEntity.builder()
                .reportFile(reportFile)
                .subject(truncate(model.getSubject(), SUBJECT_MAX, "student_results.subject"))
                .className(truncate(model.getClassName(), CLASS_NAME_MAX, "student_results.class_name"))
                .fio(truncate(model.getFio(), FIO_MAX, "student_results.fio"))
                .presence(truncate(model.getPresence(), PRESENCE_MAX, "student_results.presence"))
                .variant(truncate(model.getVariant(), VARIANT_MAX, "student_results.variant"))
                .testType(truncate(model.getTestType(), TEST_TYPE_MAX, "student_results.test_type"))
                .testDate(model.getTestDate())
                .totalScore(JsonScoreUtils.calculateTotalScore(model.getTaskScores()))
                .percentageScore(model.getPercentageScore())
                .taskScoresJson(JsonScoreUtils.mapToJson(model.getTaskScores()))
                .createdAt(LocalDateTime.now())
                .updatedAt(LocalDateTime.now())
                .schoolName(model.getSchoolName())
                .academicYear(model.getAcademicYear())
                .build();
    }

    public StudentResult toModel(StudentResultEntity entity) {
        if (entity == null) {
            return null;
        }

        StudentResult model = new StudentResult();
        model.setSubject(entity.getSubject());
        model.setClassName(entity.getClassName());
        model.setFio(entity.getFio());
        model.setPresence(entity.getPresence());
        model.setVariant(entity.getVariant());
        model.setTestType(entity.getTestType());
        model.setTestDate(entity.getTestDate());
        model.setTotalScore(entity.getTotalScore());
        model.setPercentageScore(entity.getPercentageScore()); // Добавьте это поле
        model.setTaskScores(JsonScoreUtils.jsonToMap(entity.getTaskScoresJson()));
        model.setSchoolName(entity.getSchoolName());
        model.setAcademicYear(entity.getAcademicYear());
        return model;
    }

    public void updateEntity(ReportFileEntity entity, ReportFile model) {
        if (entity == null || model == null) {
            return;
        }

        entity.setFilePath(model.getFile() != null ? model.getFile().getAbsolutePath() : entity.getFilePath());
        entity.setFileName(truncate(model.getFileName(), FILE_NAME_MAX, "report_files.file_name"));
        entity.setSubject(truncate(model.getSubject(), SUBJECT_MAX, "report_files.subject"));
        entity.setClassName(truncate(model.getClassName(), CLASS_NAME_MAX, "report_files.class_name"));
        entity.setStatus(model.getStatus());
        entity.setProcessedAt(model.getProcessedAt());
        entity.setErrorMessage(model.getErrorMessage());
        entity.setStudentCount(model.getStudentCount());
        entity.setTestDate(model.getTestDate());
        entity.setTeacher(truncate(model.getTeacher(), FIO_MAX, "report_files.teacher"));
        entity.setSchoolName(model.getSchoolName() != null ? model.getSchoolName() : entity.getSchoolName());
        entity.setTaskCount(model.getMaxScores() != null ? model.getMaxScores().size() : entity.getTaskCount());
        entity.setTestType(truncate(model.getTestType(), TEST_TYPE_MAX, "report_files.test_type"));
        entity.setComment(model.getComment());
        entity.setMaxScoresJson(JsonScoreUtils.mapToJson(model.getMaxScores()));
        entity.setUpdatedAt(LocalDateTime.now());
        entity.setAcademicYear(model.getAcademicYear());
    }

    private String truncate(String value, int maxLength, String fieldName) {
        if (value == null) {
            return null;
        }
        String trimmed = value.trim();
        if (trimmed.length() <= maxLength) {
            return trimmed;
        }
        log.warn("Поле {} превышает лимит {} символов. Значение будет обрезано.", fieldName, maxLength);
        return trimmed.substring(0, maxLength);
    }
}
