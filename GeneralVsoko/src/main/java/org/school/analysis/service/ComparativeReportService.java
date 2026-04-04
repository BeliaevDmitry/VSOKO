package org.school.analysis.service;

import java.io.File;

public interface ComparativeReportService {
    File generateEgkrEgeComparativeReport(String school, String currentAcademicYear);

    File generateEgkrEgeSubjectComparativeReport(String school, String currentAcademicYear);
}
