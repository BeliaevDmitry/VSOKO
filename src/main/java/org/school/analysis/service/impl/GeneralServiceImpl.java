package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ProcessingSummary;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.*;
import org.school.analysis.util.JsonScoreUtils;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.school.analysis.model.ProcessingStatus.*;
import static org.school.analysis.util.ValidationHelper.validateReportFile;

@Service
@RequiredArgsConstructor
@Slf4j
public class GeneralServiceImpl implements GeneralService {

    private final ParserService parserService;
    private final SavedService savedService;
    private final FileOrganizerService fileOrganizerService;
    private final AnalysisService analysisService;
    private final ExcelReportService excelReportService;

    @Override
    @Transactional
    public ProcessingSummary processAll(String folderPath) {
        ProcessingSummary summary = new ProcessingSummary();

        try {
            // 1. Найти и обработать файлы
            List<ReportFile> foundFiles = findAndProcessFiles(folderPath, summary);

            // Проверка результата
            if (foundFiles.isEmpty()) {
                log.warn("В папке {} не найдено файлов для обработки", folderPath);
                summary.setReportGenerationError("Файлы не найдены");
                return summary;
            }

            // Дополнительная проверка результатов
            validateProcessingResults(foundFiles, summary);

            // 2. Генерация отчетов
            generateReports(summary);

        } catch (Exception e) {
            log.error("Критическая ошибка при обработке отчетов", e);
            throw new RuntimeException("Ошибка обработки: " + e.getMessage(), e);
        }

        return summary;
    }

    /**
     * Валидация результатов обработки
     */
    private void validateProcessingResults(List<ReportFile> processedFiles,
                                           ProcessingSummary summary) {
        // Проверка, что хотя бы некоторые файлы обработаны успешно
        if (summary.getSuccessfullySaved() == 0) {
            log.warn("Все {} файлов обработаны с ошибками", processedFiles.size());
            summary.setReportGenerationError("Нет успешно обработанных файлов для отчетов");
        } else {
            log.info("Успешно обработано {}/{} файлов",
                    summary.getSuccessfullySaved(),
                    processedFiles.size());
        }
    }

    /**
     * Найти и обработать файлы партиями
     */
    private List<ReportFile> findAndProcessFiles(String folderPath, ProcessingSummary summary) {
        List<ReportFile> foundFiles = findReports(folderPath);
        summary.setTotalFilesFound(foundFiles.size());

        log.info("Найдено {} файлов для обработки", foundFiles.size());

        if (foundFiles.isEmpty()) {
            return foundFiles;
        }

        List<ReportFile> failedFiles = new ArrayList<>();

        // Обработка партиями
        for (int batchIndex = 0; batchIndex < foundFiles.size(); batchIndex += AppConfig.BATCH_SIZE) {
            List<ReportFile> batch = getBatch(foundFiles, batchIndex);
            List<ReportFile> processedInBatch = processBatch(batch, summary, failedFiles);

            log.debug("Партия {}-{} обработана: {} успешно",
                    batchIndex,
                    Math.min(batchIndex + AppConfig.BATCH_SIZE, foundFiles.size()),
                    processedInBatch.size());
        }

        summary.setFailedFiles(failedFiles);
        return foundFiles;
    }

    /**
     * Обработка одной партии файлов
     */
    private List<ReportFile> processBatch(List<ReportFile> batch,
                                          ProcessingSummary summary,
                                          List<ReportFile> failedFiles) {
        List<ParseResult> parseResults = parseReports(batch);

        // Статистика парсинга
        long successfullyParsed = parseResults.stream()
                .filter(ParseResult::isSuccess)
                .count();
        summary.incrementSuccessfullyParsed((int) successfullyParsed);

        // Сохранение в БД
        List<ReportFile> savedFiles = saveResultsToDatabase(parseResults);
        summary.incrementSuccessfullySaved(savedFiles.size());

        // Перемещение файлов
        List<ReportFile> movedFiles = moveProcessedFiles(savedFiles);
        summary.incrementSuccessfullyMoved(movedFiles.size());

        // Сбор информации о неудачных файлах
        collectFailedFiles(parseResults, failedFiles);

        return savedFiles;
    }

    /**
     * Получить партию файлов
     */
    private List<ReportFile> getBatch(List<ReportFile> allFiles, int startIndex) {
        int endIndex = Math.min(startIndex + AppConfig.BATCH_SIZE, allFiles.size());
        return allFiles.subList(startIndex, endIndex);
    }

    /**
     * Собрать информацию о неудачных файлах
     */
    private void collectFailedFiles(List<ParseResult> parseResults, List<ReportFile> failedFiles) {
        parseResults.stream()
                .filter(result -> !result.isSuccess() && result.getReportFile() != null)
                .forEach(result -> failedFiles.add(result.getReportFile()));
    }

    /**
     * Генерация всех отчетов
     */
    private void generateReports(ProcessingSummary summary) {
        try {
            log.info("Начинаем генерацию отчетов...");

            List<File> generatedReports = generateAllReports();

            summary.setGeneratedReportsCount(generatedReports.size());
            summary.setGeneratedReportFiles(generatedReports);

            log.info("Успешно сгенерировано {} отчетов", generatedReports.size());

        } catch (Exception e) {
            log.error("Ошибка при генерации отчетов", e);
            summary.setReportGenerationError(e.getMessage());
        }
    }

    /**
     * Генерация всех типов отчетов
     */
    private List<File> generateAllReports() {
        List<File> allReports = new ArrayList<>();

        // 1. Сводный отчет по всем тестам
        generateSummaryReport(allReports);

        // 2. Детальные отчеты по тестам
        generateTestDetailReports(allReports);

        // 3. Отчеты по учителям
        generateTeacherReports(allReports);

        return allReports;
    }

    /**
     * Генерация сводного отчета
     */
    private void generateSummaryReport(List<File> allReports) {
        List<TestSummaryDto> allTests = analysisService.getAllTestsSummary();

        if (allTests.isEmpty()) {
            log.warn("Нет данных для сводного отчета");
            return;
        }

        File summaryReport = excelReportService.generateSummaryReport(allTests);
        addReportIfValid(summaryReport, allReports, "Сводный отчет");
    }

    /**
     * Генерация детальных отчетов по тестам
     */
    private void generateTestDetailReports(List<File> allReports) {
        List<TestSummaryDto> allTests = analysisService.getAllTestsSummary();

        for (TestSummaryDto test : allTests) {
            generateSingleTestDetailReport(test, allReports);
        }
    }

    /**
     * Генерация детального отчета для одного теста
     */
    private void generateSingleTestDetailReport(TestSummaryDto test, List<File> allReports) {
        if (test.getReportFileId() == null || test.getReportFileId().trim().isEmpty()) {
            log.warn("Пропускаем тест без ID: {}", test.getFileName());
            return;
        }

        try {
            String testId = test.getReportFileId();

            log.info("Генерация детального отчета с графиками для теста: {}", test.getFileName());

            List<StudentDetailedResultDto> studentResults =
                    analysisService.getStudentDetailedResults(testId);

            Map<Integer, TaskStatisticsDto> taskStatistics =
                    analysisService.getTaskStatistics(testId);

            if (studentResults.isEmpty()) {
                log.warn("Нет данных студентов для теста: {}", test.getFileName());
                return;
            }

            if (taskStatistics == null || taskStatistics.isEmpty()) {
                log.warn("Нет статистики по заданиям для теста: {}", test.getFileName());
                return;
            }

            log.debug("Для теста {} получено: {} студентов, {} заданий",
                    test.getFileName(), studentResults.size(), taskStatistics.size());

            // Генерируем отчет с графиками на одном листе
            File detailReport = excelReportService.generateTestDetailReport(
                    test, studentResults, taskStatistics);

            if (detailReport != null && detailReport.exists()) {
                addReportIfValid(detailReport, allReports,
                        String.format("Детальный отчет с графиками для '%s'", test.getFileName()));
                log.info("✅ Отчет с графиками для теста '{}' успешно создан", test.getFileName());
            } else {
                log.warn("⚠️ Не удалось создать отчет с графиками для теста '{}'", test.getFileName());
            }

        } catch (Exception e) {
            log.error("❌ Ошибка генерации отчета с графиками для теста {}: {}",
                    test.getFileName(), e.getMessage(), e);
        }
    }

    /**
     * Генерация отчетов по учителям
     */
    private void generateTeacherReports(List<File> allReports) {
        List<String> teachers = analysisService.getAllTeachers();

        for (String teacher : teachers) {
            try {
                List<TestSummaryDto> teacherTests = analysisService.getTestsByTeacher(teacher);

                // Для каждого теста учителя получаем детальные данные
                List<TeacherTestDetailDto> teacherTestDetails = getTeacherTestDetails(teacherTests);

                // Генерируем полный отчет учителя с детальными данными
                File teacherReport = excelReportService.generateTeacherReportWithDetails(
                        teacher, teacherTests, teacherTestDetails);

                addReportIfValid(teacherReport, allReports,
                        String.format("Отчет для учителя '%s' с детализацией", teacher));

            } catch (Exception e) {
                log.error("Ошибка генерации отчета для учителя {}: {}",
                        teacher, e.getMessage(), e);
            }
        }
    }

    /**
     * Получает детальные данные для тестов учителя
     */
    /**
     * Получает детальные данные для тестов учителя
     */
    private List<TeacherTestDetailDto> getTeacherTestDetails(List<TestSummaryDto> teacherTests) {
        List<TeacherTestDetailDto> details = new ArrayList<>();

        for (TestSummaryDto test : teacherTests) {
            if (test.getReportFileId() == null || test.getReportFileId().trim().isEmpty()) {
                continue;
            }

            try {
                String testId = test.getReportFileId();

                List<StudentDetailedResultDto> studentResults =
                        analysisService.getStudentDetailedResults(testId);

                Map<Integer, TaskStatisticsDto> taskStatistics =
                        analysisService.getTaskStatistics(testId);

                TeacherTestDetailDto detailDto = TeacherTestDetailDto.builder()
                        .testSummary(test)
                        .studentResults(studentResults)
                        .taskStatistics(taskStatistics)
                        .build();

                details.add(detailDto);

            } catch (Exception e) {
                log.error("Ошибка получения детальных данных для теста {}: {}",
                        test.getFileName(), e.getMessage());
            }
        }

        return details;
    }

    /**
     * Добавить отчет в список, если он валиден
     */
    private void addReportIfValid(File report, List<File> allReports, String reportName) {
        if (report != null && report.exists()) {
            allReports.add(report);
            log.info("✅ {} сгенерирован: {}", reportName, report.getName());
        } else {
            log.warn("⚠️ {} не был сгенерирован", reportName);
        }
    }

    @Override
    @Transactional
    public List<ReportFile> saveResultsToDatabase(List<ParseResult> parseResults) {
        List<ReportFile> savedFiles = new ArrayList<>();
        AtomicInteger totalStudentsSaved = new AtomicInteger(0);

        for (ParseResult parseResult : parseResults) {
            if (!parseResult.isSuccess()) {
                continue;
            }

            ReportFile reportFile = parseResult.getReportFile();
            List<StudentResult> studentResults = parseResult.getStudentResults();

            try {
                if (processAndSaveReport(reportFile, studentResults, totalStudentsSaved)) {
                    savedFiles.add(reportFile);
                }
            } catch (Exception e) {
                handleSaveError(reportFile, e);
            }
        }

        log.info("Всего сохранено студентов: {}", totalStudentsSaved.get());
        return savedFiles;
    }

    /**
     * Обработать и сохранить один отчет
     */
    private boolean processAndSaveReport(ReportFile reportFile,
                                         List<StudentResult> studentResults,
                                         AtomicInteger totalStudentsSaved) {
        if (!validateReportFile(reportFile)) {
            markFileAsInvalid(reportFile, "Некорректные данные");
            return false;
        }

        enrichReportFileData(reportFile);
        calculateStudentScores(studentResults, reportFile);

        int savedCount = savedService.saveAll(reportFile, studentResults);

        if (savedCount > 0) {
            markFileAsSaved(reportFile, savedCount);
            totalStudentsSaved.addAndGet(savedCount);
            return true;
        } else {
            markFileAsSaveFailed(reportFile, "Не удалось сохранить данные");
            return false;
        }
    }

    /**
     * Обогатить данные отчета
     */
    private void enrichReportFileData(ReportFile reportFile) {
        if (reportFile.getMaxScores() != null) {
            reportFile.setTaskCount(reportFile.getMaxScores().size());
        }
    }

    /**
     * Рассчитать баллы студентов
     */
    private void calculateStudentScores(List<StudentResult> studentResults,
                                        ReportFile reportFile) {
        studentResults.forEach(student -> {
            if (student.getTaskScores() != null) {
                int totalScore = JsonScoreUtils.calculateTotalScore(student.getTaskScores());
                student.setTotalScore(totalScore);

                if (reportFile.getMaxTotalScore() > 0) {
                    double percentage = (totalScore * 100.0) / reportFile.getMaxTotalScore();
                    student.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                }
            }
        });
    }

    /**
     * Отметить файл как сохраненный
     */
    private void markFileAsSaved(ReportFile reportFile, int savedCount) {
        reportFile.setStatus(SAVED);
        log.info("✅ Файл '{}' сохранен ({} студентов, {} заданий)",
                reportFile.getFileName(), savedCount, reportFile.getTaskCount());
    }

    /**
     * Отметить файл как невалидный
     */
    private void markFileAsInvalid(ReportFile reportFile, String errorMessage) {
        reportFile.setStatus(ERROR_SAVING);
        reportFile.setErrorMessage(errorMessage);
        log.warn("⚠️ Файл '{}' содержит некорректные данные", reportFile.getFileName());
    }

    /**
     * Отметить файл как неудачно сохраненный
     */
    private void markFileAsSaveFailed(ReportFile reportFile, String errorMessage) {
        reportFile.setStatus(ERROR_SAVING);
        reportFile.setErrorMessage(errorMessage);
        log.warn("⚠️ Файл '{}' не сохранен (0 студентов)", reportFile.getFileName());
    }

    /**
     * Обработать ошибку сохранения
     */
    private void handleSaveError(ReportFile reportFile, Exception e) {
        reportFile.setStatus(ERROR_SAVING);
        reportFile.setErrorMessage("Ошибка БД: " + e.getMessage());
        log.error("❌ Ошибка сохранения файла '{}': {}",
                reportFile.getFileName(), e.getMessage());
    }

    @Override
    public List<ReportFile> findReports(String folderPath) {
        return fileOrganizerService.findReportFiles(folderPath);
    }

    @Override
    public List<ParseResult> parseReports(List<ReportFile> reportFiles) {
        return parserService.parseFiles(reportFiles);
    }

    @Override
    public List<ReportFile> moveProcessedFiles(List<ReportFile> successfullyProcessedFiles) {
        return fileOrganizerService.moveFilesToSubjectFolders(successfullyProcessedFiles);
    }
}