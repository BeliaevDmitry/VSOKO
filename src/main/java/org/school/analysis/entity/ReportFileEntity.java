package org.school.analysis.entity;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.UpdateTimestamp;
import org.school.analysis.model.ProcessingStatus;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@Entity
@Table(name = "report_files")
@Getter
@Setter
@ToString(exclude = {"studentResults"})  // Теперь исключаем только studentResults
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ReportFileEntity {

    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private UUID id;

    @Column(name = "file_path", nullable = false, length = 500)
    private String filePath;

    @Column(name = "file_name", nullable = false, length = 255)
    private String fileName;

    @Column(name = "file_hash", nullable = false, length = 64, unique = true)
    private String fileHash;

    @Column(nullable = false, length = 100)
    private String subject;

    @Column(name = "class_name", nullable = false, length = 50)
    private String className;

    @Enumerated(EnumType.STRING)
    @Column(nullable = false, length = 50)
    private ProcessingStatus status;

    @Column(name = "processed_at")
    private LocalDateTime processedAt;

    @Column(name = "error_message", columnDefinition = "TEXT")
    private String errorMessage;

    @Column(name = "student_count")
    private Integer studentCount = 0;

    @Column(name = "test_date")
    private LocalDate testDate;

    @Column(length = 200)
    private String teacher;

    @Builder.Default
    @Column(length = 200)
    private String school = "ГБОУ №7";

    @Column(name = "task_count")
    private Integer taskCount;

    @Column(name = "test_type", length = 50)
    private String testType;

    @Column(columnDefinition = "TEXT")
    private String comment;

    @Column(name = "max_scores_json", columnDefinition = "TEXT")
    private String maxScoresJson;

    @CreationTimestamp
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @UpdateTimestamp
    @Column(name = "updated_at", nullable = false)
    private LocalDateTime updatedAt;

    @OneToMany(mappedBy = "reportFile", cascade = CascadeType.ALL)
    @Builder.Default
    private List<StudentResultEntity> studentResults = new ArrayList<>();
}