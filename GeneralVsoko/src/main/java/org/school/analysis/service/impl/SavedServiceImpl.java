package org.school.analysis.service.impl;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.model.entity.ReportFileEntity;
import org.school.analysis.model.entity.StudentResultEntity;
import org.school.analysis.mapper.ReportMapper;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.repository.ReportFileRepository;
import org.school.analysis.repository.StudentResultRepository;
import org.school.analysis.service.SavedService;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.File;
import java.nio.file.Files;
import java.security.MessageDigest;
import java.util.*;

@Service
@RequiredArgsConstructor
@Slf4j
public class SavedServiceImpl implements SavedService {

    @PersistenceContext
    private final EntityManager entityManager; // Добавляем EntityManager

    private final ReportFileRepository reportFileRepository;
    private final StudentResultRepository studentResultRepository;
    private final ReportMapper reportMapper;

    @Override
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

            // 2. Создаем ReportFileEntity (без ID - будет сгенерирован БД)
            ReportFileEntity reportFileEntity = reportMapper.toEntity(reportFile);
            reportFileEntity.setFileHash(fileHash);
            reportFileEntity.setStudentCount(studentResults.size());

            // 3. Используем entityManager.persist() для новой сущности файла
            entityManager.persist(reportFileEntity);
            entityManager.flush(); // Получаем сгенерированный ID

            log.info("Файл сохранен с ID: {}", reportFileEntity.getId());

            // 4. Подготавливаем студентов
            Map<Integer, Integer> maxScoresMap = reportFile.getMaxScores();
            int maxTotalScore = calculateMaxTotalScore(maxScoresMap);

            List<StudentResultEntity> studentEntities = new ArrayList<>();

            for (StudentResult student : studentResults) {
                StudentResultEntity entity = reportMapper.toEntity(student, reportFileEntity);

                if (entity.getTaskScoresJson() != null && maxTotalScore > 0) {
                    Map<Integer, Integer> taskScores = JsonScoreUtils.jsonToMap(entity.getTaskScoresJson());
                    if (taskScores != null && !taskScores.isEmpty()) {
                        int totalScore = JsonScoreUtils.calculateTotalScore(taskScores);
                        double percentage = (totalScore * 100.0) / maxTotalScore;
                        entity.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                    }
                }

                studentEntities.add(entity);
            }

            // 5. Сохраняем студентов через entityManager.persist()
            for (StudentResultEntity entity : studentEntities) {
                entityManager.persist(entity);
            }

            entityManager.flush();

            log.info("✅ Сохранено {} студентов из файла {}",
                    studentEntities.size(), reportFile.getFileName());
            return studentEntities.size();

        } catch (Exception e) {
            log.error("❌ Ошибка сохранения файла {}: {}",
                    reportFile.getFileName(), e.getMessage(), e);
            throw new RuntimeException("Ошибка сохранения в базу данных: " + e.getMessage(), e);
        }
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

    @Override
    @Transactional(readOnly = true)
    public ReportFile getReportFileById(UUID id) {
        ReportFileEntity entity = reportFileRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Файл не найден с ID: " + id));
        return reportMapper.toModel(entity);
    }

    @Override
    @Transactional(readOnly = true)
    public List<StudentResult> getStudentResultsByReportFileId(UUID reportFileId) {
        List<StudentResultEntity> entities = studentResultRepository
                .findByReportFileId(reportFileId);

        return entities.stream()
                .map(reportMapper::toModel)
                .toList();
    }

    @Override
    @Transactional(readOnly = true)
    public long countByReportFileId(UUID reportFileId) {
        return studentResultRepository.countByReportFileId(reportFileId);
    }
}