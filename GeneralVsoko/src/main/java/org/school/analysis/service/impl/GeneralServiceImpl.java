package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.model.dto.StudentDetailedResultDto;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.school.analysis.model.dto.TeacherTestDetailDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.service.*;
import org.school.analysis.util.JsonScoreUtils;
import org.school.analysis.util.PerformanceTracker;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.io.File;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import static org.school.analysis.config.AppConfig.*;
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
    private final TeacherService teacherService;

    @Override

    public void processAll() {
        PerformanceTracker.startProgram();
        int totalFilesFound = 0;
        int totalSuccessfullySaved = 0;
        int totalGeneratedReports = 0;

        try {
            for (String school : SCHOOLS) {
                PerformanceTracker.SchoolProcessingMetrics schoolMetrics =
                        PerformanceTracker.startSchoolProcessing(school);

                String currentAcademicYear = ALL_ACADEMIC_YEAR.get(0);
                System.out.println("  Учебный год: " + currentAcademicYear);
                System.out.println("  Обработка школы " + school);

                initTeacherDatabase(school);

                String folderPath = INPUT_FOLDER.replace("{школа}", school);
                List<ReportFile> failedFiles = new ArrayList<>();
                int[] schoolCounters = new int[3]; // [0]=найдено, [1]=сохранено, [2]=отчеты

                try {
                    long phase1Start = System.currentTimeMillis();
                    List<ReportFile> foundFiles = findAndProcessFiles(folderPath, failedFiles, schoolCounters);
                    long phase1Time = System.currentTimeMillis() - phase1Start;
                    PerformanceTracker.recordPhaseTime(school, "fileFinding", Duration.ofMillis(phase1Time));

                    long phase2Start = System.currentTimeMillis();
                    if (!foundFiles.isEmpty()) {
                        validateProcessingResults(foundFiles, schoolCounters[1]);
                        schoolCounters[2] = generateReports(school, currentAcademicYear);
                    }
                    long phase2Time = System.currentTimeMillis() - phase2Start;
                    PerformanceTracker.recordPhaseTime(school, "reportGeneration", Duration.ofMillis(phase2Time));

                    PerformanceTracker.recordPhaseTime(school, "fileProcessing",
                            Duration.ofMillis(phase1Time + phase2Time));

                    PerformanceTracker.finishSchoolProcessing(
                            schoolMetrics,
                            schoolCounters[0],
                            schoolCounters[1],
                            schoolCounters[2]
                    );

                    printSchoolSummary(school, schoolCounters[1], schoolCounters[2], failedFiles);

                    totalFilesFound += schoolCounters[0];
                    totalSuccessfullySaved += schoolCounters[1];
                    totalGeneratedReports += schoolCounters[2];

                } catch (Exception e) {
                    log.error("❌ Критическая ошибка при обработке школы {}: {}", school, e.getMessage(), e);
                    PerformanceTracker.finishSchoolProcessing(
                            schoolMetrics,
                            schoolCounters[0],
                            schoolCounters[1],
                            schoolCounters[2]
                    );
                }
            }

            printSummary(totalFilesFound, totalSuccessfullySaved, totalGeneratedReports);

            System.out.println("\n" + "=".repeat(80));
            System.out.println("СТАТИСТИКА ОБРАБОТКИ");
            System.out.println("=".repeat(80));

            String schoolsStats = PerformanceTracker.getSchoolsStatistics();
            System.out.println(schoolsStats);

            String finalSummary = PerformanceTracker.getFinalSummary();
            System.out.println(finalSummary);

            System.out.println("=".repeat(80));
            System.out.println("ОБРАБОТКА ЗАВЕРШЕНА!");

        } finally {
            PerformanceTracker.clear();
        }
    }

    @Override
    @Transactional(propagation = Propagation.REQUIRES_NEW, noRollbackFor = {Exception.class})
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
                log.error("Ошибка сохранения отчета {}: {}",
                        reportFile.getFileName(), e.getMessage());
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
        try {
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
        } catch (Exception e) {
            log.error("Ошибка обработки отчета {}: {}",
                    reportFile.getFileName(), e.getMessage());
            markFileAsSaveFailed(reportFile, "Ошибка обработки: " + e.getMessage());
            return false;
        }
    }

    /**
     * Инициализация базы данных учителей
     */
    private void initTeacherDatabase(String school) {
        log.info("=== ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ УЧИТЕЛЕЙ ===");

        try {
            // Загружаем статистику до импорта
            Map<String, Object> statsBefore = teacherService.getStatistics();
            log.info("До импорта: {} активных учителей в БД", statsBefore.get("activeTeachers"));

            // Импортируем из файла (если есть изменения)
            teacherService.importTeachersFromExcel(school);

            // Загружаем статистику после импорта
            Map<String, Object> statsAfter = teacherService.getStatistics();
            log.info("После импорта: {} активных учителей в БД", statsAfter.get("activeTeachers"));

        } catch (Exception e) {
            log.error("Ошибка при инициализации базы учителей: {}", e.getMessage());
            log.warn("Продолжаем обработку без проверки учителей");
        }

        log.info("=== ЗАВЕРШЕНО ИНИЦИАЛИЗАЦИЯ УЧИТЕЛЕЙ ===\n");
    }

    /**
     * Валидация результатов обработки
     */
    private void validateProcessingResults(List<ReportFile> processedFiles, int successfullySaved) {
        // Проверка, что хотя бы некоторые файлы обработаны успешно
        if (successfullySaved == 0) {
            log.warn("Все {} файлов обработаны с ошибками", processedFiles.size());
        } else {
            log.info("Успешно обработано {}/{} файлов",
                    successfullySaved,
                    processedFiles.size());
        }
    }

    /**
     * Найти и обработать файлы партиями
     */
    private List<ReportFile> findAndProcessFiles(String folderPath,
                                                 List<ReportFile> failedFiles,
                                                 int[] schoolCounters) {
        List<ReportFile> foundFiles = findReports(folderPath);
        schoolCounters[0] = foundFiles.size();

        log.info("Найдено {} файлов для обработки", foundFiles.size());

        if (foundFiles.isEmpty()) {
            return foundFiles;
        }

        // Обработка партиями
        for (int batchIndex = 0; batchIndex < foundFiles.size(); batchIndex += AppConfig.BATCH_SIZE) {
            List<ReportFile> batch = getBatch(foundFiles, batchIndex);
            List<ReportFile> processedInBatch = processBatch(batch, failedFiles, schoolCounters);

            log.debug("Партия {}-{} обработана: {} успешно",
                    batchIndex,
                    Math.min(batchIndex + AppConfig.BATCH_SIZE, foundFiles.size()),
                    processedInBatch.size());
        }

        return foundFiles;
    }

    /**
     * Обработка одной партии файлов
     */
    private List<ReportFile> processBatch(List<ReportFile> batch,
                                          List<ReportFile> failedFiles,
                                          int[] schoolCounters) {
        List<ParseResult> parseResults = parseReports(batch);

        // Сохранение в БД
        List<ReportFile> savedFiles = saveResultsToDatabase(parseResults);
        schoolCounters[1] += savedFiles.size();
        moveProcessedFiles(savedFiles);

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
    private int generateReports(String school, String currentAcademicYear) {
        try {
            log.info("Начинаем генерацию отчетов...");

            List<File> generatedReports = generateAllReports(school, currentAcademicYear);

            log.info("Успешно сгенерировано {} отчетов", generatedReports.size());
            return generatedReports.size();

        } catch (Exception e) {
            log.error("Ошибка при генерации отчетов", e);
            return 0;
        }
    }

    /**
     * Генерация всех типов отчетов
     */
    private List<File> generateAllReports(String school, String currentAcademicYear) {
        List<File> allReports = new ArrayList<>();

        // 1. Сводный отчет по всем тестам
        generateSummaryReport(allReports, school, currentAcademicYear);

        // 2. Детальные отчеты по тестам
        generateTestDetailReports(allReports, school, currentAcademicYear);

        // 3. Отчеты по учителям
        generateTeacherReports(allReports, school, currentAcademicYear);

        return allReports;
    }

    /**
     * Генерация сводного отчета
     */
    private void generateSummaryReport(List<File> allReports,
                                       String schoolName,
                                       String currentAcademicYear) {
        List<TestSummaryDto> allTests = analysisService.getAllTestsSummary(schoolName, currentAcademicYear);

        if (allTests.isEmpty()) {
            log.warn("Нет данных для сводного отчета");
            return;
        }

        File summaryReport = excelReportService.generateSummaryReport(allTests, schoolName);
        addReportIfValid(summaryReport, allReports, "Сводный отчет");
    }

    /**
     * Генерация детальных отчетов по тестам
     */
    private void generateTestDetailReports(List<File> allReports, String schoolName,
                                           String currentAcademicYear) {
        List<TestSummaryDto> allTests = analysisService.getAllTestsSummary(schoolName, currentAcademicYear);

        for (TestSummaryDto test : allTests) {
            generateSingleTestDetailReport(test, allReports, schoolName);
        }
    }

    /**
     * Генерация детального отчета для одного теста
     */
    private void generateSingleTestDetailReport(TestSummaryDto test, List<File> allReports,
                                                String schoolName) {
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
                    test, studentResults, taskStatistics, schoolName);

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
    private void generateTeacherReports(List<File> allReports, String schoolName,
                                        String currentAcademicYear) {
        List<String> teachers = analysisService.getAllTeachers(schoolName, currentAcademicYear);
        log.info("✅ размер teachers '{}' ", teachers.size());
        for (String teacher : teachers) {
            try {
                log.info("✅ зашли в  generateTeacherReports и анализируем '{}' ", teacher);
                List<TestSummaryDto> teacherTests = analysisService.getTestsByTeacher(teacher,
                        schoolName, currentAcademicYear);
                log.info("✅ размер teacherTests '{}' ", teacherTests.size());
                // Для каждого теста учителя получаем детальные данные
                List<TeacherTestDetailDto> teacherTestDetails = getTeacherTestDetails(teacherTests);
                log.info("✅ размер teacherTestDetails '{}' ", teacherTestDetails.size());
                // Генерируем полный отчет учителя с детальными данными
                File teacherReport = excelReportService.generateTeacherReportWithDetails(
                        teacher, teacherTests, teacherTestDetails, schoolName);

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

    /**
     * Вывод промежуточных результатов по школе
     */
    private void printSchoolSummary(String schoolName,
                                    int processedFiles,
                                    int generatedReportsCount,
                                    List<ReportFile> failedFiles) {
        log.info("\n" + "-".repeat(60));
        log.info("✅ ИТОГИ ОБРАБОТКИ ШКОЛЫ: {}", schoolName);
        log.info("-".repeat(60));
        log.info("✅ Обработано файлов: {}", processedFiles);
        log.info("📊 Создано отчетов: {}", generatedReportsCount);
        log.info("❌ Ошибок обработки: {}", failedFiles.size());

        for (ReportFile failedFile : failedFiles) {
            String reason = failedFile.getErrorMessage() != null
                    ? failedFile.getErrorMessage()
                    : String.valueOf(failedFile.getStatus());
            log.warn("  • {} — {}", failedFile.getFileName(), reason);
        }
        log.info("=".repeat(60));
    }
    
    private static void printSummary(int totalFilesFound,
                                     int totalSuccessfullySaved,
                                     int totalGeneratedReports) {
        System.out.println("\n" + "=".repeat(50));
        System.out.println("ИТОГИ ОБРАБОТКИ:");
        System.out.println("=".repeat(50));
        System.out.println("📁 Всего найдено файлов: " + totalFilesFound);
        System.out.println("✅ Обработано файлов: " + totalSuccessfullySaved);
        System.out.println("📊 Создано отчетов: " + totalGeneratedReports);
        System.out.println("❌ Не обработано файлов: " + (totalFilesFound - totalSuccessfullySaved));
        System.out.println("\n" + "=".repeat(50));
    }
}
