package org.school.analysis.repository;

import org.school.analysis.model.Teacher;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface TeacherRepository extends JpaRepository<Teacher, Long> {

    Optional<Teacher> findByFullName(String fullName);

    Optional<Teacher> findByNormalizedFullName(String normalizedFullName);

    List<Teacher> findByLastNameIgnoreCase(String lastName);

    @Query("SELECT t FROM Teacher t WHERE t.normalizedFullName LIKE %:searchTerm% OR t.fullName LIKE %:searchTerm%")
    List<Teacher> searchByNormalizedName(@Param("searchTerm") String searchTerm);

    List<Teacher> findByIsActiveTrue();

    List<Teacher> findByIsActiveFalse();

    @Query("SELECT COUNT(t) FROM Teacher t WHERE t.isActive = true")
    long countActiveTeachers();

    List<Teacher> findByLastNameContainingIgnoreCase(String lastNamePart);
}