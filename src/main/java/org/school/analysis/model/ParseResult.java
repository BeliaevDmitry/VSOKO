package org.school.analysis.model;

import lombok.Data;
import java.time.LocalDateTime;
import java.util.List;

@Data
public class ParseResult {
    // Обязательные поля
    private ReportFile reportFile;
    private List<StudentResult> studentResults;
    private boolean success;
    private String errorMessage;
    private int parsedStudents;

    // Опциональные поля (можно удалить если не нужны)
    private TestMetadata metadata;
    private LocalDateTime parsedAt;
    private String parserVersion = "1.0";

    /**
     * Конструктор для успешного парсинга
     */
    public static ParseResult success(ReportFile reportFile,
                                      List<StudentResult> studentResults) {
        ParseResult result = new ParseResult();
        result.setReportFile(reportFile);
        result.setStudentResults(studentResults);
        result.setSuccess(true);
        result.setParsedStudents(studentResults.size());
        result.setParsedAt(LocalDateTime.now());
        return result;
    }

    /**
     * Конструктор для неудачного парсинга
     */
    public static ParseResult error(ReportFile reportFile, String errorMessage) {
        ParseResult result = new ParseResult();
        result.setReportFile(reportFile);
        result.setSuccess(false);
        result.setErrorMessage(errorMessage);
        result.setParsedAt(LocalDateTime.now());
        return result;
    }
}