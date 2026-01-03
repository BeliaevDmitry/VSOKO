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
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

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

            // 2. Создаем ReportFileEntity и сохраняем сразу
            ReportFileEntity reportFileEntity = createReportFileEntity(reportFile, fileHash, studentResults.size());
            entityManager.persist(reportFileEntity);
            entityManager.flush(); // Получаем ID для связей

            // 3. Сохраняем максимальные баллы пакетно
            saveMaxScoresBatch(reportFile, reportFileEntity);

            // 4. Сохраняем студентов пакетно
            AtomicInteger savedCount = new AtomicInteger(0);
            List<StudentTaskScoreEntity> allTaskScores = new ArrayList<>();
            saveStudentsBatch(studentResults, reportFileEntity, savedCount, allTaskScores);

            // 5. Сохраняем все баллы заданий пакетно
            saveTaskScoresBatch(allTaskScores);

            // 6. Обновляем счетчик студентов
            reportFileEntity.setStudentCount(savedCount.get());

            log.info("Сохранено {} студентов из файла {}", savedCount.get(), reportFile.getFileName());
            return savedCount.get();

        } catch (Exception e) {
            log.error("Ошибка сохранения файла {}: {}",
                    reportFile.getFileName(), e.getMessage(), e);
            return 0;
        }
    }

    private ReportFileEntity createReportFileEntity(ReportFile reportFile, String fileHash, int studentCount) {
        ReportFileEntity entity = ReportFileEntity.builder()
                .filePath(reportFile.getFile().getAbsolutePath())
                .fileName(reportFile.getFile().getName())
                .fileHash(fileHash)
                .subject(reportFile.getSubject())
                .className(reportFile.getClassName())
                .status(reportFile.getStatus())
                .processedAt(LocalDateTime.now())
                .errorMessage(reportFile.getErrorMessage())
                .studentCount(studentCount)
                .testDate(reportFile.getTestDate())
                .teacher(reportFile.getTeacher())
                .school(reportFile.getSchool())
                .taskCount(reportFile.getTaskCount())
                .maxTotalScore(reportFile.getMaxTotalScore())
                .testType(reportFile.getTestType())
                .comment(reportFile.getComment())
                .build();
        return entity;
    }

    private void saveMaxScoresBatch(ReportFile reportFile, ReportFileEntity reportFileEntity) {
        if (reportFile.getMaxScores() == null || reportFile.getMaxScores().isEmpty()) {
            return;
        }

        List<MaxScoreEntity> maxScores = new ArrayList<>();
        reportFile.getMaxScores().forEach((taskNumber, maxScore) -> {
            MaxScoreEntity maxScoreEntity = MaxScoreEntity.builder()
                    .reportFile(reportFileEntity)
                    .taskNumber(taskNumber)
                    .maxScore(maxScore)
                    .build();
            maxScores.add(maxScoreEntity);
            reportFileEntity.getMaxScores().add(maxScoreEntity);
        });

        // Пакетное сохранение максимальных баллов
        for (int i = 0; i < maxScores.size(); i++) {
            entityManager.persist(maxScores.get(i));
            if (i % 50 == 0 && i > 0) {
                entityManager.flush();
                entityManager.clear();
            }
        }
        entityManager.flush();
        entityManager.clear();
    }

    private void saveStudentsBatch(List<StudentResult> studentResults,
                                   ReportFileEntity reportFileEntity,
                                   AtomicInteger savedCount,
                                   List<StudentTaskScoreEntity> allTaskScores) {

        // Собираем все максимальные баллы в Map для быстрого доступа
        Map<Integer, Integer> maxScoresMap = new HashMap<>();
        if (reportFileEntity.getMaxScores() != null) {
            for (MaxScoreEntity maxScoreEntity : reportFileEntity.getMaxScores()) {
                maxScoresMap.put(maxScoreEntity.getTaskNumber(), maxScoreEntity.getMaxScore());
            }
        }

        for (int i = 0; i < studentResults.size(); i++) {
            StudentResult student = studentResults.get(i);

            try {
                StudentResultEntity studentEntity = createStudentEntity(student, reportFileEntity, maxScoresMap);
                entityManager.persist(studentEntity);

                // Собираем баллы заданий для пакетного сохранения
                if (student.getTaskScores() != null) {
                    collectTaskScores(student, studentEntity, maxScoresMap, allTaskScores);
                }

                savedCount.incrementAndGet();

                // Периодически сбрасываем в базу (оптимизация памяти)
                if (i % 50 == 0 && i > 0) {
                    entityManager.flush();
                    entityManager.clear();
                    // После clear нужно повторно привязать reportFileEntity
                    reportFileEntity = entityManager.merge(reportFileEntity);
                }

            } catch (Exception e) {
                log.error("Ошибка сохранения студента {}: {}", student.getFio(), e.getMessage());
            }
        }

        // Финальный flush
        entityManager.flush();
        entityManager.clear();
    }

    private StudentResultEntity createStudentEntity(StudentResult student,
                                                    ReportFileEntity reportFileEntity,
                                                    Map<Integer, Integer> maxScoresMap) {
        StudentResultEntity entity = StudentResultEntity.builder()
                .reportFile(reportFileEntity)
                .subject(student.getSubject())
                .className(student.getClassName())
                .fio(student.getFio())
                .presence(student.getPresence())
                .variant(student.getVariant())
                .testType(student.getTestType())
                .testDate(student.getTestDate())
                .build();

        // Вычисляем общий балл
        if (student.getTaskScores() != null && !student.getTaskScores().isEmpty()) {
            int totalScore = student.getTaskScores().values().stream()
                    .mapToInt(Integer::intValue)
                    .sum();
            entity.setTotalScore(totalScore);

            // Процент выполнения
            if (reportFileEntity.getMaxTotalScore() != null &&
                    reportFileEntity.getMaxTotalScore() > 0) {
                double percentage = (totalScore * 100.0) / reportFileEntity.getMaxTotalScore();
                entity.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
            }
        }

        return entity;
    }

    private void collectTaskScores(StudentResult student,
                                   StudentResultEntity studentEntity,
                                   Map<Integer, Integer> maxScoresMap,
                                   List<StudentTaskScoreEntity> allTaskScores) {

        student.getTaskScores().forEach((taskNumber, score) -> {
            Integer maxScore = maxScoresMap.getOrDefault(taskNumber, 0);
            StudentTaskScoreEntity taskScore = StudentTaskScoreEntity.builder()
                    .studentResult(studentEntity)
                    .taskNumber(taskNumber)
                    .score(score)
                    .maxScore(maxScore)
                    .build();
            allTaskScores.add(taskScore);
        });
    }

    private void saveTaskScoresBatch(List<StudentTaskScoreEntity> taskScores) {
        if (taskScores.isEmpty()) {
            return;
        }

        for (int i = 0; i < taskScores.size(); i++) {
            entityManager.persist(taskScores.get(i));
            if (i % 100 == 0 && i > 0) { // Больший размер для мелких записей
                entityManager.flush();
                entityManager.clear();
            }
        }
        entityManager.flush();
        entityManager.clear();
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