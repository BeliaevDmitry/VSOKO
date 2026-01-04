package org.school.analysis.model.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TestSummaryDto {
    private String school;
    private String subject;
    private String className;
    private LocalDate testDate;
    private String testType;
    private String teacher;
    private Integer studentsCount;      // Кол-во учеников писавших
    private Integer classSize;          // Кол-во учеников в классе
    private Integer taskCount;          // Кол-во заданий теста
    private Integer maxTotalScore;      // Макс. балл
    private Double averageScore;        // Средний балл теста
    private String fileName;            // Имя исходного файла
    private Integer studentsTotal;      // Всего учеников в классе (из справочника)
    private Integer studentsPresent;    // Присутствовало на тесте
    private Integer studentsAbsent;     // Отсутствовало на тесте

    // Вычисляемое поле для процента присутствия
    public Double getAttendancePercentage() {
        if (studentsTotal == null || studentsTotal == 0) return 0.0;
        return Math.round((studentsPresent * 100.0 / studentsTotal) * 10.0) / 10.0;
    }

}