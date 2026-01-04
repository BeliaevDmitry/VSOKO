package org.school.analysis.repository.impl;

import jakarta.persistence.EntityManager;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.entity.StudentResultEntity;
import org.school.analysis.mapper.ReportMapper;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.repository.ReportFileRepository;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import java.io.File;
import java.nio.file.Files;
import java.security.MessageDigest;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

@Repository
@RequiredArgsConstructor
@Slf4j
public class StudentResultRepositoryImpl {

    private final EntityManager entityManager;
    private final ReportFileRepository reportFileRepository;
    private final ReportMapper reportMapper; // Добавляем маппер

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

            // 2. Создаем ReportFileEntity через маппер
            ReportFileEntity reportFileEntity = reportMapper.toEntity(reportFile);
            reportFileEntity.setFileHash(fileHash); // Устанавливаем хэш
            reportFileEntity.setStudentCount(studentResults.size()); // Устанавливаем количество студентов

            entityManager.persist(reportFileEntity);
            entityManager.flush(); // Получаем ID для связей

            // 3. Сохраняем студентов пакетно с использованием маппера
            AtomicInteger savedCount = new AtomicInteger(0);
            saveStudentsBatch(studentResults, reportFileEntity, savedCount);

            log.info("Сохранено {} студентов из файла {}", savedCount.get(), reportFile.getFileName());
            return savedCount.get();

        } catch (Exception e) {
            log.error("Ошибка сохранения файла {}: {}",
                    reportFile.getFileName(), e.getMessage(), e);
            return 0;
        }
    }

    private void saveStudentsBatch(List<StudentResult> studentResults,
                                   ReportFileEntity reportFileEntity,
                                   AtomicInteger savedCount) {

        // Получаем максимальные баллы из JSON для расчета процентов
        Map<Integer, Integer> maxScoresMap = JsonScoreUtils.jsonToMap(reportFileEntity.getMaxScoresJson());
        int maxTotalScore = calculateMaxTotalScore(maxScoresMap);

        for (int i = 0; i < studentResults.size(); i++) {
            StudentResult student = studentResults.get(i);

            try {
                // Создаем сущность через маппер
                StudentResultEntity studentEntity = reportMapper.toEntity(student, reportFileEntity);

                // Дополнительные расчеты (процент выполнения)
                if (studentEntity.getTaskScoresJson() != null && maxTotalScore > 0) {
                    Map<Integer, Integer> taskScores = JsonScoreUtils.jsonToMap(studentEntity.getTaskScoresJson());
                    if (taskScores != null && !taskScores.isEmpty()) {
                        int totalScore = JsonScoreUtils.calculateTotalScore(taskScores);
                        double percentage = (totalScore * 100.0) / maxTotalScore;
                        studentEntity.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                    }
                }

                entityManager.persist(studentEntity);
                savedCount.incrementAndGet();

                // Периодически сбрасываем в базу (оптимизация памяти)
                if (i % 50 == 0 && i > 0) {
                    entityManager.flush();
                    entityManager.clear();
                    // После clear нужно повторно привязать reportFileEntity
                    reportFileEntity = entityManager.merge(reportFileEntity);
                }

            } catch (Exception e) {
                log.error("Ошибка сохранения студента {}: {}", student.getFio(), e.getMessage(), e);
            }
        }

        // Финальный flush
        entityManager.flush();
        entityManager.clear();
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
     * Упрощенный метод с использованием маппера
     */
    @Transactional
    public int saveAllWithMapper(ReportFile reportFile, List<StudentResult> studentResults) {
        try {
            // 1. Проверка дубликата
            String fileHash = calculateFileHash(reportFile.getFile());
            if (reportFileRepository.existsByFileHash(fileHash)) {
                log.warn("Файл уже был обработан: {}", reportFile.getFileName());
                return 0;
            }

            // 2. Создаем ReportFileEntity через маппер
            ReportFileEntity reportFileEntity = reportMapper.toEntity(reportFile);
            reportFileEntity.setFileHash(fileHash);
            reportFileEntity.setStudentCount(studentResults.size());

            // 3. Создаем StudentResultEntity для каждого студента через маппер
            List<StudentResultEntity> studentEntities = new ArrayList<>();

            // Получаем максимальные баллы для расчета процентов
            Map<Integer, Integer> maxScoresMap = reportFile.getMaxScores();
            int maxTotalScore = calculateMaxTotalScore(maxScoresMap);

            for (StudentResult student : studentResults) {
                StudentResultEntity studentEntity = reportMapper.toEntity(student, reportFileEntity);

                // Рассчитываем процент выполнения
                if (studentEntity.getTaskScoresJson() != null && maxTotalScore > 0) {
                    Map<Integer, Integer> taskScores = JsonScoreUtils.jsonToMap(studentEntity.getTaskScoresJson());
                    if (taskScores != null && !taskScores.isEmpty()) {
                        int totalScore = JsonScoreUtils.calculateTotalScore(taskScores);
                        double percentage = (totalScore * 100.0) / maxTotalScore;
                        studentEntity.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                    }
                }

                studentEntities.add(studentEntity);
            }

            // 4. Сохраняем все одной транзакцией
            entityManager.persist(reportFileEntity);

            for (StudentResultEntity studentEntity : studentEntities) {
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

    /**
     * Метод для получения ReportFile по ID с использованием маппера
     */
    public ReportFile getReportFileById(UUID id) {
        ReportFileEntity entity = entityManager.find(ReportFileEntity.class, id);
        return reportMapper.toModel(entity);
    }

    /**
     * Метод для получения результатов студента по ID файла
     */
    public List<StudentResult> getStudentResultsByReportFileId(UUID reportFileId) {
        String query = "SELECT sr FROM StudentResultEntity sr WHERE sr.reportFile.id = :reportFileId";
        List<StudentResultEntity> entities = entityManager.createQuery(query, StudentResultEntity.class)
                .setParameter("reportFileId", reportFileId)
                .getResultList();

        List<StudentResult> results = new ArrayList<>();
        for (StudentResultEntity entity : entities) {
            results.add(reportMapper.toModel(entity));
        }
        return results;
    }
}