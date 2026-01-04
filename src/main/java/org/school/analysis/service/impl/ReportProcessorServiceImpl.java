package org.school.analysis.service.impl;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.*;

import lombok.extern.slf4j.Slf4j;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ProcessingSummary;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.dto.StudentTestResultDto;
import org.school.analysis.model.dto.TestResultsDto;
import org.school.analysis.model.dto.TestSummaryDto;
import org.school.analysis.repository.impl.StudentResultRepositoryImpl;
import org.school.analysis.service.*;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import static org.school.analysis.config.AppConfig.FINAL_REPORT_FOLDER;
import static org.school.analysis.model.ProcessingStatus.*;
import static org.school.analysis.util.DateTimeFormatters.DISPLAY_DATE;

@Service
@RequiredArgsConstructor
@Slf4j
public class ReportProcessorServiceImpl implements ReportProcessorService {

    private final ReportParserService parserService;
    private final StudentResultRepositoryImpl repositoryImpl;
    private final FileOrganizerService fileOrganizerService;
    private final AnalysisService analysisService;
    private final ExcelReportService excelReportService;


    private String finalReportFolder = FINAL_REPORT_FOLDER;


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

                log.debug("Обработано файлов: {}-{} из {}", i, end, foundFiles.size());
            }

            // 3. Генерация отчетов после обработки всех файлов
            generateReportsAfterProcessing(summary);

            // 4. Итоговая статистика
            log.info("ИТОГО: Найдено={}, Распарсено={}, Сохранено={}, Перемещено={}, Отчеты={}",
                    summary.getTotalFilesFound(),
                    summary.getSuccessfullyParsed(),
                    summary.getSuccessfullySaved(),
                    summary.getSuccessfullyMoved(),
                    summary.getGeneratedReportsCount());

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
            int detailedReportsCount = generateDetailedTeacherReports(allTests);
            log.info("Сгенерировано {} детальных отчетов по учителям", detailedReportsCount);

        } catch (Exception e) {
            log.error("Ошибка при генерации сводных отчетов: {}", e.getMessage(), e);
            throw e;
        }

        return allGeneratedFiles;
    }

    /**
     * Генерация детальных отчетов для каждого учителя/теста
     */
    private int generateDetailedTeacherReports(List<TestSummaryDto> allTests) {
        int generatedCount = 0;

        for (TestSummaryDto test : allTests) {
            try {
                // Получаем детальные результаты теста
                TestResultsDto testResults = analysisService.getTestResults(
                        test.getSubject(),
                        test.getClassName(),
                        test.getTestDate(),
                        test.getTeacher()
                );

                if (testResults != null &&
                        testResults.getStudentResults() != null &&
                        !testResults.getStudentResults().isEmpty()) {

                    // Генерируем отчет для учителя с диаграммой
                    File teacherReport = excelReportService.generateTeacherReport(
                            testResults,
                            true // включаем диаграмму
                    );

                    if (teacherReport != null && teacherReport.exists()) {
                        generatedCount++;
                        log.debug("Сгенерирован отчет для {} {} {}: {}",
                                test.getSubject(), test.getClassName(),
                                test.getTestDate(), teacherReport.getName());
                    }
                }

            } catch (Exception e) {
                log.error("Ошибка при генерации отчета для теста {} {} {}: {}",
                        test.getSubject(), test.getClassName(), test.getTestDate(),
                        e.getMessage());
                // Продолжаем с другими тестами
            }
        }

        return generatedCount;
    }

    /**
     * Альтернативный метод: генерация отчетов с группировкой по учителям
     */
    private List<File> generateReportsGroupedByTeacher(List<TestSummaryDto> allTests) {
        List<File> generatedFiles = new ArrayList<>();

        try {
            // Группируем тесты по учителю
            Map<String, List<TestSummaryDto>> testsByTeacher = new HashMap<>();

            for (TestSummaryDto test : allTests) {
                String teacher = test.getTeacher();
                if (teacher != null && !teacher.trim().isEmpty()) {
                    testsByTeacher
                            .computeIfAbsent(teacher, k -> new ArrayList<>())
                            .add(test);
                }
            }

            log.info("Найдено {} учителей для генерации отчетов", testsByTeacher.size());

            // Генерируем сводный отчет для каждого учителя
            for (Map.Entry<String, List<TestSummaryDto>> entry : testsByTeacher.entrySet()) {
                String teacher = entry.getKey();
                List<TestSummaryDto> teacherTests = entry.getValue();

                try {
                    // Сортируем тесты по дате
                    teacherTests.sort(Comparator.comparing(TestSummaryDto::getTestDate));

                    // Генерируем отчет по учителю
                    File teacherSummaryReport = generateTeacherSummaryReport(teacher, teacherTests);
                    if (teacherSummaryReport != null && teacherSummaryReport.exists()) {
                        generatedFiles.add(teacherSummaryReport);
                        log.info("Сводный отчет для учителя {}: {}", teacher, teacherSummaryReport.getName());
                    }

                } catch (Exception e) {
                    log.error("Ошибка при генерации отчета для учителя {}: {}", teacher, e.getMessage());
                }
            }

        } catch (Exception e) {
            log.error("Ошибка при группировке отчетов по учителям: {}", e.getMessage(), e);
        }

        return generatedFiles;
    }

    /**
     * Генерация сводного отчета по конкретному учителю
     */
    private File generateTeacherSummaryReport(String teacher, List<TestSummaryDto> teacherTests) {
        try {
            // Создаем workbook
            XSSFWorkbook workbook = new XSSFWorkbook();

            // Лист с данными
            XSSFSheet sheet = workbook.createSheet("Отчет по учителю");

            // Заголовок
            Row headerRow = sheet.createRow(0);
            Cell headerCell = headerRow.createCell(0);
            headerCell.setCellValue("Отчет по тестам учителя: " + teacher);

            // Заголовки таблицы
            String[] headers = {"Предмет", "Класс", "Дата теста", "Тип теста",
                    "Кол-во учеников", "Средний балл", "Статус"};

            Row tableHeaderRow = sheet.createRow(2);
            for (int i = 0; i < headers.length; i++) {
                tableHeaderRow.createCell(i).setCellValue(headers[i]);
            }

            // Заполняем данными
            int rowNum = 3;
            for (TestSummaryDto test : teacherTests) {
                Row row = sheet.createRow(rowNum++);

                // Получаем детальные результаты для расчета статистики
                TestResultsDto results = analysisService.getTestResults(
                        test.getSubject(),
                        test.getClassName(),
                        test.getTestDate(),
                        test.getTeacher()
                );

                double averageScore = 0.0;
                if (results != null && results.getStudentResults() != null) {
                    averageScore = results.getStudentResults().stream()
                            .filter(s -> s.getTotalScore() != null)
                            .mapToInt(StudentTestResultDto::getTotalScore)
                            .average()
                            .orElse(0.0);
                }

                row.createCell(0).setCellValue(test.getSubject());
                row.createCell(1).setCellValue(test.getClassName());
                row.createCell(2).setCellValue(test.getTestDate().format(DISPLAY_DATE));
                row.createCell(3).setCellValue(test.getTestType() != null ? test.getTestType() : "");
                               row.createCell(5).setCellValue(averageScore);
                row.createCell(6).setCellValue("Обработан");
            }

            // Сохраняем файл
            String fileName = String.format("teacher_report_%s_%s.xlsx",
                    teacher.replaceAll("[^a-zA-Zа-яА-Я0-9]", "_"),
                    LocalDate.now().format(DISPLAY_DATE));

            Path reportsPath = Paths.get(finalReportFolder);
            if (!Files.exists(reportsPath)) {
                Files.createDirectories(reportsPath);
            }

            Path filePath = reportsPath.resolve(fileName);
            try (FileOutputStream outputStream = new FileOutputStream(filePath.toFile())) {
                workbook.write(outputStream);
            }

            workbook.close();

            return filePath.toFile();

        } catch (Exception e) {
            log.error("Ошибка при создании отчета для учителя {}: {}", teacher, e.getMessage());
            return null;
        }
    }

    @Override
    @Transactional
    public List<ReportFile> saveResultsToDatabase(List<ParseResult> parseResults) {
        List<ReportFile> savedFiles = new ArrayList<>();

        for (ParseResult parseResult : parseResults) {
            if (!parseResult.isSuccess()) {
                continue;
            }

            ReportFile reportFile = parseResult.getReportFile();

            try {
                // СОХРАНЯЕМ В БД через единый репозиторий!
                int savedCount = repositoryImpl.saveAll(
                        reportFile,
                        parseResult.getStudentResults()
                );

                if (savedCount > 0) {
                    reportFile.setStatus(SAVED);
                    savedFiles.add(reportFile);
                    log.info("✅ Файл {} сохранен ({} студентов)",
                            reportFile.getFileName(), savedCount);
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
                        reportFile.getFileName(), e.getMessage());
            }
        }

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