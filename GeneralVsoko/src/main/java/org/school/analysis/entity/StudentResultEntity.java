package org.school.analysis.entity;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.UpdateTimestamp;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.UUID;

@Entity
@Table(name = "student_results")
@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StudentResultEntity {

    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private UUID id;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "report_file_id", nullable = false)
    @ToString.Exclude
    private ReportFileEntity reportFile;

    @Column(name = "subject", nullable = false, length = 100)
    private String subject;

    @Column(name = "class_name", nullable = false, length = 50)
    private String className;

    @Column(name = "fio", nullable = false, length = 200)
    private String fio;

    @Column(name = "presence", nullable = true, length = 50)
    private String presence;

    @Column(name = "variant", length = 100)
    private String variant;

    @Column(name = "test_type", length = 50)
    private String testType;

    @Column(name = "test_date", nullable = false)
    private LocalDate testDate;

    @Column(name = "total_score")
    private Integer totalScore;

    @Column(name = "percentage_score")
    private Double percentageScore;

    // Баллы храним как JSON
    @Column(name = "task_scores_json", columnDefinition = "TEXT")
    private String taskScoresJson;

    @CreationTimestamp
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @UpdateTimestamp
    @Column(name = "updated_at", nullable = false)
    private LocalDateTime updatedAt;

    @Builder.Default
    @Column(length = 200)
    private String ACADEMIC_YEAR = "2025-2026";

    @Builder.Default
    @Column(length = 200)
    private String school = "ГБОУ №7";
}