package org.school.analysis.repository.impl;

import jakarta.persistence.EntityManager;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.entity.*;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.repository.ReportFileRepository;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import java.io.File;
import java.nio.file.Files;
import java.security.MessageDigest;
import java.util.List;
import java.util.UUID;

@Repository
@RequiredArgsConstructor
@Slf4j
public class StudentResultRepositoryImpl {

    private final EntityManager entityManager;
    private final ReportFileRepository reportFileRepository;

    @Transactional
    public int saveAll(ReportFile reportFile, List<StudentResult> studentResults) {
        log.info("Сохранение файла {} с {} студентами",
                reportFile.getFileName(), studentResults.size());

        try {
            // 1. Проверка дубликата
            String fileHash = calculateFileHash(reportFile.getFile());
            if (reportFileRepository.existsByFileHash(fileHash)) {
                log.warn("Файл уже был обработан: {}", reportFile.getFileName());
                return 0;
            }

            // 2. Создаем ReportFileEntity
            ReportFileEntity reportFileEntity = ReportFileEntity.builder()
                    .filePath(reportFile.getFile().getAbsolutePath())
                    .fileName(reportFile.getFile().getName())
                    .fileHash(fileHash)
                    .subject(reportFile.getSubject())
                    .className(reportFile.getClassName())
                    .status(reportFile.getStatus())
                    .processedAt(reportFile.getProcessedAt())
                    .errorMessage(reportFile.getErrorMessage())
                    .studentCount(studentResults.size())
                    .testDate(reportFile.getTestDate())
                    .teacher(reportFile.getTeacher())
                    .school(reportFile.getSchool())
                    .taskCount(reportFile.getTaskCount())
                    .maxTotalScore(reportFile.getMaxTotalScore())
                    .testType(reportFile.getTestType())
                    .comment(reportFile.getComment())
                    .build();

            // 3. Добавляем максимальные баллы
            if (reportFile.getMaxScores() != null) {
                reportFile.getMaxScores().forEach((taskNumber, maxScore) -> {
                    reportFileEntity.addMaxScore(taskNumber, maxScore);
                });
            }

            // 4. Сохраняем файл
            entityManager.persist(reportFileEntity);

            // 5. Сохраняем студентов
            int savedCount = 0;
            for (StudentResult student : studentResults) {
                try {
                    StudentResultEntity studentEntity = StudentResultEntity.builder()
                            .reportFile(reportFileEntity)
                            .subject(student.getSubject())
                            .className(student.getClassName())
                            .fio(student.getFio())
                            .presence(student.getPresence())
                            .variant(student.getVariant())
                            .testType(student.getTestType())
                            .testDate(student.getTestDate())
                            .build();

                    // Добавляем баллы по заданиям
                    if (student.getTaskScores() != null) {
                        student.getTaskScores().forEach((taskNumber, score) -> {
                            // Находим максимальный балл для задания
                            Integer maxScore = reportFileEntity.getMaxScores().stream()
                                    .filter(ms -> ms.getTaskNumber().equals(taskNumber))
                                    .map(MaxScoreEntity::getMaxScore)
                                    .findFirst()
                                    .orElse(0);

                            studentEntity.addTaskScore(taskNumber, score, maxScore);
                        });
                    }

                    // Вычисляем общий балл
                    if (student.getTaskScores() != null) {
                        int totalScore = student.getTaskScores().values().stream()
                                .mapToInt(Integer::intValue)
                                .sum();
                        studentEntity.setTotalScore(totalScore);

                        // Процент выполнения
                        if (reportFileEntity.getMaxTotalScore() != null &&
                                reportFileEntity.getMaxTotalScore() > 0) {
                            double percentage = (totalScore * 100.0) / reportFileEntity.getMaxTotalScore();
                            studentEntity.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                        }
                    }

                    entityManager.persist(studentEntity);
                    savedCount++;

                } catch (Exception e) {
                    log.error("Ошибка сохранения студента {}: {}",
                            student.getFio(), e.getMessage());
                }
            }

            // 6. Обновляем счетчик
            reportFileEntity.setStudentCount(savedCount);

            log.info("Сохранено {} студентов из файла {}", savedCount, reportFile.getFileName());
            return savedCount;

        } catch (Exception e) {
            log.error("Ошибка сохранения файла {}: {}",
                    reportFile.getFileName(), e.getMessage(), e);
            return 0;
        }
    }

    private String calculateFileHash(File file) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            byte[] fileBytes = Files.readAllBytes(file.toPath());
            byte[] hashBytes = digest.digest(fileBytes);

            StringBuilder hexString = new StringBuilder();
            for (byte b : hashBytes) {
                String hex = Integer.toHexString(0xff & b);
                if (hex.length() == 1) hexString.append('0');
                hexString.append(hex);
            }
            return hexString.toString();
        } catch (Exception e) {
            log.error("Ошибка вычисления хеша файла: {}", e.getMessage());
            return UUID.randomUUID().toString().replace("-", "");
        }
    }
}