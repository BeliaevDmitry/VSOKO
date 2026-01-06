package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.config.AppConfig;
import org.school.analysis.entity.ReportFileEntity;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ProcessingSummary;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
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
public class ReportProcessorServiceImpl implements ReportProcessorService {

    private final ReportParserService parserService;
    private final StudentResultService studentResultService;
    private final FileOrganizerService fileOrganizerService;
    private final AnalysisService analysisService;
    private final ExcelReportService excelReportService;

    @Override
    @Transactional
    public ProcessingSummary processAll(String folderPath) {
        ProcessingSummary summary = new ProcessingSummary();
        List<ReportFile> failedFiles = new ArrayList<>();

        try {
            // 1. Найти файлы
            List<ReportFile> foundFiles = findReports(folderPath);

            summary.setTotalFilesFound(foundFiles.size());
            log.info("Найдено {} файлов", foundFiles.size());

            // 2. Обработка файлов небольшими партиями (для оптимизации памяти)
            for (int i = 0; i < foundFiles.size(); i += AppConfig.BATCH_SIZE) {
                int end = Math.min(i + AppConfig.BATCH_SIZE, foundFiles.size());
                List<ReportFile> batch = foundFiles.subList(i, end);

                // Обработка партии
                List<ParseResult> parseResults = parseReports(batch);
                List<ReportFile> savedFiles = saveResultsToDatabase(parseResults);
                List<ReportFile> movedFiles = moveProcessedFiles(savedFiles);

                // Обновление статистики
                summary.setSuccessfullyParsed(summary.getSuccessfullyParsed() +
                        (int) parseResults.stream().filter(ParseResult::isSuccess).count());
                summary.setSuccessfullySaved(summary.getSuccessfullySaved() + savedFiles.size());
                summary.setSuccessfullyMoved(summary.getSuccessfullyMoved() + movedFiles.size());

                // Собираем информацию о неудачных файлах
                for (ParseResult parseResult : parseResults) {
                    if (!parseResult.isSuccess() && parseResult.getReportFile() != null) {
                        failedFiles.add(parseResult.getReportFile());
                    }
                }

                log.debug("Обработано файлов: {}-{} из {}", i, end, foundFiles.size());
            }

            // 3. Сохраняем информацию о неудачных файлах
            summary.setFailedFiles(failedFiles);

            // 4. Генерация отчетов после обработки всех файлов
            generateReportsAfterProcessing(summary);

        } catch (Exception e) {
            log.error("Ошибка при обработке отчетов: {}", e.getMessage(), e);
            throw new RuntimeException("Ошибка при обработке отчетов: " + e.getMessage(), e);
        }

        return summary;
    }

    /**
     * Генерация отчетов после обработки файлов
     */
    private void generateReportsAfterProcessing(ProcessingSummary summary) {
        try {
            log.info("Начинаем генерацию отчетов...");

            // 1. Генерация сводного отчета по всем тестам
            List<File> generatedReports = generateSummaryReports();

            summary.setGeneratedReportsCount(generatedReports.size());
            summary.setGeneratedReportFiles(generatedReports);

            log.info("Сгенерировано {} отчетов", generatedReports.size());

        } catch (Exception e) {
            log.error("Ошибка при генерации отчетов: {}", e.getMessage(), e);
            // Не прерываем основной процесс из-за ошибок в отчетах
            summary.setReportGenerationError(e.getMessage());
        }
    }

    /**
     * Генерация сводного отчета и отчетов по учителям
     */
    // В методе generateSummaryReports() добавьте генерацию отчетов по учителям:

    private List<File> generateSummaryReports() {
        List<File> allGeneratedFiles = new ArrayList<>();

        try {
            // 1. Получаем список всех тестов
            List<TestSummaryDto> allTests = analysisService.getAllTestsSummary();

            if (allTests.isEmpty()) {
                log.warn("Нет данных для генерации отчетов");
                return allGeneratedFiles;
            }

            log.info("Найдено {} тестов для генерации отчетов", allTests.size());

            // 2. Генерируем сводный отчет по всем тестам
            File summaryReport = excelReportService.generateSummaryReport(allTests);
            if (summaryReport != null && summaryReport.exists()) {
                allGeneratedFiles.add(summaryReport);
                log.info("Сводный отчет сгенерирован: {}", summaryReport.getName());
            }

            // 3. Генерируем детальные отчеты для каждого теста
            generateDetailedTestReports(allTests, allGeneratedFiles);

            // 4. Генерируем отчеты по учителям
            generateTeacherReports(allGeneratedFiles);

        } catch (Exception e) {
            log.error("Ошибка при генерации сводных отчетов: {}", e.getMessage(), e);
            throw e;
        }

        return allGeneratedFiles;
    }

    /**
     * Генерация детальных отчетов для каждого теста
     */
    private void generateDetailedTestReports(List<TestSummaryDto> allTests, List<File> allGeneratedFiles) {
        for (TestSummaryDto test : allTests) {
            try {
                // Проверяем наличие ID
                if (test.getReportFileId() == null || test.getReportFileId().trim().isEmpty()) {
                    log.warn("Нет ID для файла: {}", test.getFileName());
                    continue;
                }

                log.info("Генерация детального отчета для теста ID: {}, {}",
                        test.getReportFileId(), test.getFileName());

                // Получаем детальные данные напрямую по ID
                List<StudentDetailedResultDto> studentResults =
                        analysisService.getStudentDetailedResults(test.getReportFileId());

                Map<Integer, TaskStatisticsDto> taskStatistics =
                        analysisService.getTaskStatistics(test.getReportFileId());

                if (studentResults.isEmpty()) {
                    log.warn("Нет данных студентов для теста: {}", test.getFileName());
                    continue;
                }

                // Генерируем отчет
                File detailReport = excelReportService.generateTestDetailReport(
                        test, studentResults, taskStatistics);

                if (detailReport != null && detailReport.exists()) {
                    allGeneratedFiles.add(detailReport);
                    log.info("✅ Детальный отчет для теста {} сгенерирован: {}",
                            test.getFileName(), detailReport.getName());
                }
            } catch (Exception e) {
                log.error("❌ Ошибка при генерации детального отчета для теста {}: {}",
                        test.getFileName(), e.getMessage(), e);
            }
        }
    }

    /**
     * Генерация отчетов по учителям
     */
    private void generateTeacherReports(List<File> allGeneratedFiles) {
        List<String> teachers = analysisService.getAllTeachers();

        for (String teacher : teachers) {
            try {
                List<TestSummaryDto> teacherTests = analysisService.getTestsByTeacher(teacher);

                File teacherReport = excelReportService.generateTeacherReport(teacher, teacherTests);
                if (teacherReport != null && teacherReport.exists()) {
                    allGeneratedFiles.add(teacherReport);
                    log.info("Отчет для учителя {} сгенерирован", teacher);
                }
            } catch (Exception e) {
                log.error("Ошибка при генерации отчета для учителя {}: {}", teacher, e.getMessage());
            }
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
                // Валидация данных перед сохранением
                if (!validateReportFile(reportFile)) {
                    reportFile.setStatus(ERROR_SAVING);
                    reportFile.setErrorMessage("Некорректные данные");
                    log.warn("⚠️ Файл {} содержит некорректные данные", reportFile.getFileName());
                    continue;
                }

                // Устанавливаем taskCount
                if (reportFile.getMaxScores() != null) {
                    reportFile.setTaskCount(reportFile.getMaxScores().size());
                }

                // Вычисляем totalScore для каждого студента
                for (StudentResult student : studentResults) {
                    if (student.getTaskScores() != null) {
                        int totalScore = JsonScoreUtils.calculateTotalScore(student.getTaskScores());
                        student.setTotalScore(totalScore);

                        if (reportFile.getMaxTotalScore() > 0) {
                            double percentage = (totalScore * 100.0) / reportFile.getMaxTotalScore();
                            student.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                        }
                    }
                }

                // 🔄 ИЗМЕНЕНИЕ: используем новый сервис
                int savedCount = studentResultService.saveAll(reportFile, studentResults);

                if (savedCount > 0) {
                    reportFile.setStatus(SAVED);
                    savedFiles.add(reportFile);
                    totalStudentsSaved.addAndGet(savedCount);
                    log.info("✅ Файл {} сохранен ({} студентов, {} заданий)",
                            reportFile.getFileName(),
                            savedCount,
                            reportFile.getTaskCount());
                } else {
                    reportFile.setStatus(ERROR_SAVING);
                    reportFile.setErrorMessage("Не удалось сохранить данные в БД");
                    log.warn("⚠️ Файл {} не сохранен (0 студентов)",
                            reportFile.getFileName());
                }

            } catch (Exception e) {
                reportFile.setStatus(ERROR_SAVING);
                reportFile.setErrorMessage("Ошибка БД: " + e.getMessage());
                log.error("❌ Ошибка сохранения файла {}: {}",
                        reportFile.getFileName(), e.getMessage(), e);
            }
        }

        log.info("Всего сохранено студентов: {}", totalStudentsSaved.get());
        return savedFiles;
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