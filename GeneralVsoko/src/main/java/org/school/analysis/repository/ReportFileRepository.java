package org.school.analysis.repository;

import org.school.analysis.model.entity.ReportFileEntity;
import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.JpaSpecificationExecutor;
import org.springframework.stereotype.Repository;

import java.util.List;
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

    Optional<ReportFileEntity> findByFileName(String fileName);

    List<ReportFileEntity> findByTeacher(String teacher, Sort sort);

    List<ReportFileEntity> findBySchoolNameAndAcademicYear(
            String schoolName,
            String academicYear,
            Sort sort
    );


    /**
     * Найти тесты по учителю, школе и учебному году
     */
    List<ReportFileEntity> findByTeacherAndAcademicYear(
            String schoolName,
            String academicYear
    );

    /**
     * Найти тесты по учителю, школе и учебному году
     */
    List<ReportFileEntity> findByTeacherAndSchoolNameAndAcademicYear(
            String teacher,
            String schoolName,
            String academicYear,
            Sort sort
    );

    /**
     * Найти тесты по учителю и школе
     */
    List<ReportFileEntity> findByTeacherAndSchoolName(
            String teacher,
            String schoolName,
            Sort sort
    );
}