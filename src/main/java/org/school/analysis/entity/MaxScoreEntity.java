package org.school.analysis.entity;

import jakarta.persistence.*;
import lombok.*;

import java.util.UUID;

@Entity
@Table(name = "report_file_max_scores")
@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
@Builder
@EqualsAndHashCode(onlyExplicitlyIncluded = true)
public class MaxScoreEntity {

    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    @EqualsAndHashCode.Include
    private UUID id;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "report_file_id", nullable = false)
    @ToString.Exclude
    private ReportFileEntity reportFile;

    @Column(name = "task_number", nullable = false)
    private Integer taskNumber;

    @Column(name = "max_score", nullable = false)
    private Integer maxScore;
}