package org.school.analysis.mapper;

import org.school.analysis.model.entity.ReportFileEntity;
import org.school.analysis.model.entity.StudentResultEntity;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Component;

import java.time.LocalDateTime;

@Component
public class ReportMapper {

    public ReportFileEntity toEntity(ReportFile model) {
        if (model == null) {
            return null;
        }

        return ReportFileEntity.builder()
                .filePath(model.getFile() != null ? model.getFile().getAbsolutePath() : "")
                .fileName(model.getFileName())
                .fileHash("") // Вычисляется отдельно
                .subject(model.getSubject())
                .className(model.getClassName())
                .status(model.getStatus())
                .processedAt(model.getProcessedAt())
                .errorMessage(model.getErrorMessage())
                .studentCount(model.getStudentCount())
                .testDate(model.getTestDate())
                .teacher(model.getTeacher())
                .schoolName(model.getSchoolName() != null ? model.getSchoolName() : "ГБОУ №7")
                .academicYear(model.getAcademicYear() != null ? model.getAcademicYear() : "2025-2026")
                .taskCount(model.getMaxScores() != null ? model.getMaxScores().size() : 0)
                .testType(model.getTestType())
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
                .subject(model.getSubject())
                .className(model.getClassName())
                .fio(model.getFio())
                .presence(model.getPresence())
                .variant(model.getVariant())
                .testType(model.getTestType())
                .testDate(model.getTestDate())
                .totalScore(JsonScoreUtils.calculateTotalScore(model.getTaskScores()))
                .percentageScore(model.getPercentageScore()) // Добавьте это поле
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
        entity.setFileName(model.getFileName());
        entity.setSubject(model.getSubject());
        entity.setClassName(model.getClassName());
        entity.setStatus(model.getStatus());
        entity.setProcessedAt(model.getProcessedAt());
        entity.setErrorMessage(model.getErrorMessage());
        entity.setStudentCount(model.getStudentCount());
        entity.setTestDate(model.getTestDate());
        entity.setTeacher(model.getTeacher());
        entity.setSchoolName(model.getSchoolName() != null ? model.getSchoolName() : entity.getSchoolName());
        entity.setTaskCount(model.getMaxScores() != null ? model.getMaxScores().size() : entity.getTaskCount());
        entity.setTestType(model.getTestType());
        entity.setComment(model.getComment());
        entity.setMaxScoresJson(JsonScoreUtils.mapToJson(model.getMaxScores()));
        entity.setUpdatedAt(LocalDateTime.now());
        entity.setAcademicYear(model.getAcademicYear());
    }
}