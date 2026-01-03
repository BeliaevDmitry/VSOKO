package org.school.analysis.repository;

import org.school.analysis.entity.StudentResultEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.UUID;

@Repository
public interface StudentResultRepository extends JpaRepository<StudentResultEntity, UUID> {
}