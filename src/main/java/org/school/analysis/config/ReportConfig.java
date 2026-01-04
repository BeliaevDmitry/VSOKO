package org.school.analysis.config;

import lombok.Getter;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;

@Configuration
@Getter
public class ReportConfig {

    @Value("${report.detailed.enabled:true}")
    private boolean detailedReportsEnabled;

    @Value("${report.teacher.enabled:true}")
    private boolean teacherReportsEnabled;

    @Value("${report.max.students.per.sheet:100}")
    private int maxStudentsPerSheet;

    @Value("${report.chart.width:600}")
    private int chartWidth;

    @Value("${report.chart.height:400}")
    private int chartHeight;
}