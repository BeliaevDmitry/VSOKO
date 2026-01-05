package org.school.analysis.repository;

import org.school.analysis.entity.StudentResultEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.JpaSpecificationExecutor;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.UUID;

@Repository
public interface StudentResultRepository extends
        JpaRepository<StudentResultEntity, UUID>,
        JpaSpecificationExecutor<StudentResultEntity> {

    /**
     * Найти все результаты по ID файла отчета
     */
    List<StudentResultEntity> findByReportFileId(UUID reportFileId);

    /**
     * Подсчитать количество результатов по ID файла отчета
     */
    long countByReportFileId(UUID reportFileId);

    /**
     * Найти результаты с процентом выполнения выше указанного
     */
    List<StudentResultEntity> findByReportFileIdAndPercentageScoreGreaterThanEqual(
            UUID reportFileId, Double minPercentage);

    /**
     * Удалить все результаты по ID файла отчета
     */
    void deleteByReportFileId(UUID reportFileId);
}