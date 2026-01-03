package org.school.analysis.entity;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.UpdateTimestamp;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@Entity
@Table(name = "student_results")
@Getter
@Setter
@ToString(exclude = {"reportFile", "taskScores"})
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class StudentResultEntity {

    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private UUID id;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "report_file_id", nullable = false)
    private ReportFileEntity reportFile;

    @Column(nullable = false, length = 100)
    private String subject;

    @Column(name = "class_name", nullable = false, length = 50)
    private String className;

    @Column(nullable = false, length = 200)
    private String fio;

    @Column(nullable = false, length = 50)
    private String presence;

    @Column(length = 100)
    private String variant;

    @Column(name = "test_type", length = 50)
    private String testType;

    @Column(name = "test_date", nullable = false)
    private LocalDate testDate;

    @Column(name = "total_score")
    private Integer totalScore;

    // ИСПРАВЛЕНО: убрали precision и scale для Double
    @Column(name = "percentage_score")
    private Double percentageScore;

    @CreationTimestamp
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @UpdateTimestamp
    @Column(name = "updated_at", nullable = false)
    private LocalDateTime updatedAt;

    // Связь с баллами по заданиям
    @OneToMany(mappedBy = "studentResult", cascade = CascadeType.ALL, orphanRemoval = true)
    @Builder.Default
    private List<StudentTaskScoreEntity> taskScores = new ArrayList<>();

    // Вспомогательные методы
    public void addTaskScore(Integer taskNumber, Integer score, Integer maxScore) {
        StudentTaskScoreEntity taskScore = StudentTaskScoreEntity.builder()
                .studentResult(this)
                .taskNumber(taskNumber)
                .score(score)
                .maxScore(maxScore)
                .build();
        taskScores.add(taskScore);
    }

    // Вычисляемые поля (не сохраняются в БД)
    @Transient
    public boolean wasPresent() {
        return "Был".equalsIgnoreCase(presence);
    }

    @Transient
    public Integer getScoreForTask(Integer taskNumber) {
        return taskScores.stream()
                .filter(ts -> ts.getTaskNumber().equals(taskNumber))
                .map(StudentTaskScoreEntity::getScore)
                .findFirst()
                .orElse(null);
    }
}