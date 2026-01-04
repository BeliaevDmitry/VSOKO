package org.school.analysis.repository.impl;

import jakarta.persistence.EntityManager;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.repository.ReportFileRepository;
import org.school.analysis.util.JsonScoreUtils;
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

            // 3. Сохраняем студентов пакетно
            AtomicInteger savedCount = new AtomicInteger(0);
            saveStudentsBatch(studentResults, reportFileEntity, savedCount);

            // 4. Обновляем счетчик студентов
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
                // Максимальные баллы сохраняем как JSON
                .maxScoresJson(JsonScoreUtils.mapToJson(reportFile.getMaxScores()))
                .testType(reportFile.getTestType())
                .comment(reportFile.getComment())
                .build();
        return entity;
    }

    private void saveStudentsBatch(List<StudentResult> studentResults,
                                   ReportFileEntity reportFileEntity,
                                   AtomicInteger savedCount) {

        // Получаем максимальные баллы из JSON
        Map<Integer, Integer> maxScoresMap = JsonScoreUtils.jsonToMap(reportFileEntity.getMaxScoresJson());
        int maxTotalScore = calculateMaxTotalScore(maxScoresMap);

        for (int i = 0; i < studentResults.size(); i++) {
            StudentResult student = studentResults.get(i);

            try {
                StudentResultEntity studentEntity = createStudentEntity(
                        student,
                        reportFileEntity,
                        maxScoresMap,
                        maxTotalScore
                );
                entityManager.persist(studentEntity);

                savedCount.incrementAndGet();

                // Периодически сбрасываем в базу (оптимизация памяти)
                if (i % 50 == 0 && i > 0) {
                    entityManager.flush();
                    entityManager.clear();
                    // После clear нужно повторно привязать reportFileEntity
                    reportFileEntity = entityManager.merge(reportFileEntity);
                    // Также нужно обновить maxScoresMap
                    maxScoresMap = JsonScoreUtils.jsonToMap(reportFileEntity.getMaxScoresJson());
                    maxTotalScore = calculateMaxTotalScore(maxScoresMap);
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
                                                    Map<Integer, Integer> maxScoresMap,
                                                    int maxTotalScore) {
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

        // Сохраняем баллы как JSON
        if (student.getTaskScores() != null && !student.getTaskScores().isEmpty()) {
            // Сохраняем JSON
            entity.setTaskScoresJson(JsonScoreUtils.mapToJson(student.getTaskScores()));

            // Вычисляем общий балл
            int totalScore = JsonScoreUtils.calculateTotalScore(student.getTaskScores());
            entity.setTotalScore(totalScore);

            // Процент выполнения
            if (maxTotalScore > 0) {
                double percentage = (totalScore * 100.0) / maxTotalScore;
                entity.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
            }
        }

        return entity;
    }

    private int calculateMaxTotalScore(Map<Integer, Integer> maxScoresMap) {
        if (maxScoresMap == null || maxScoresMap.isEmpty()) {
            return 0;
        }
        return maxScoresMap.values().stream()
                .mapToInt(Integer::intValue)
                .sum();
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

    /**
     * Альтернативный метод: массовое сохранение через JPA репозиторий
     */
    @Transactional
    public int saveAllOptimized(ReportFile reportFile, List<StudentResult> studentResults) {
        try {
            // 1. Проверка дубликата
            String fileHash = calculateFileHash(reportFile.getFile());
            if (reportFileRepository.existsByFileHash(fileHash)) {
                log.warn("Файл уже был обработан: {}", reportFile.getFileName());
                return 0;
            }

            // 2. Создаем ReportFileEntity
            ReportFileEntity reportFileEntity = createReportFileEntity(reportFile, fileHash, studentResults.size());

            // 3. Создаем StudentResultEntity для каждого студента
            Map<Integer, Integer> maxScoresMap = reportFile.getMaxScores();
            int maxTotalScore = calculateMaxTotalScore(maxScoresMap);

            List<StudentResultEntity> studentEntities = new ArrayList<>();
            for (StudentResult student : studentResults) {
                StudentResultEntity studentEntity = createStudentEntity(
                        student,
                        reportFileEntity,
                        maxScoresMap,
                        maxTotalScore
                );
                studentEntities.add(studentEntity);
            }

            // 4. Сохраняем все одной транзакцией
            reportFileEntity = entityManager.merge(reportFileEntity);

            for (StudentResultEntity studentEntity : studentEntities) {
                studentEntity.setReportFile(reportFileEntity);
                entityManager.persist(studentEntity);
            }

            entityManager.flush();

            log.info("Сохранено {} студентов из файла {}", studentEntities.size(), reportFile.getFileName());
            return studentEntities.size();

        } catch (Exception e) {
            log.error("Ошибка сохранения файла {}: {}",
                    reportFile.getFileName(), e.getMessage(), e);
            return 0;
        }
    }
}