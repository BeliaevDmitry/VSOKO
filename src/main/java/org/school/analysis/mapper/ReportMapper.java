package org.school.analysis.mapper;

import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Component;

import java.time.LocalDateTime;
import java.util.UUID;

@Component
public class ReportMapper {

    public ReportFileEntity toEntity(ReportFile model) {
        if (model == null) {
            return null;
        }

        return ReportFileEntity.builder()
                .id(UUID.randomUUID())
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
                .school(model.getSchool() != null ? model.getSchool() : "ГБОУ №7")
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
        model.setSchool(entity.getSchool());
        model.setTaskCount(entity.getTaskCount());
        model.setMaxScores(JsonScoreUtils.jsonToMap(entity.getMaxScoresJson()));
        model.setTestType(entity.getTestType());
        model.setComment(entity.getComment());
        return model;
    }

    public StudentResultEntity toEntity(StudentResult model, ReportFileEntity reportFile) {
        if (model == null) {
            return null;
        }

        return StudentResultEntity.builder()
                .id(UUID.randomUUID())
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
        entity.setSchool(model.getSchool() != null ? model.getSchool() : entity.getSchool());
        entity.setTaskCount(model.getMaxScores() != null ? model.getMaxScores().size() : entity.getTaskCount());
        entity.setTestType(model.getTestType());
        entity.setComment(model.getComment());
        entity.setMaxScoresJson(JsonScoreUtils.mapToJson(model.getMaxScores()));
        entity.setUpdatedAt(LocalDateTime.now());
    }
}