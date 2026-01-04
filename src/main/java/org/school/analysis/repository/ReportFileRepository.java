package org.school.analysis.repository;

import org.school.analysis.entity.ReportFileEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.JpaSpecificationExecutor;
import org.springframework.stereotype.Repository;

import java.util.Optional;
import java.util.UUID;

@Repository
public interface ReportFileRepository extends
        JpaRepository<ReportFileEntity, UUID>,
        JpaSpecificationExecutor<ReportFileEntity> {

    // Используется в StudentResultRepositoryImpl
    boolean existsByFileHash(String fileHash);

    // Опционально, может пригодиться
    Optional<ReportFileEntity> findByFileHash(String fileHash);
}