package org.school.analysis.service.impl;

import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ReportFile;
import org.school.analysis.repository.impl.StudentResultRepositoryImpl;
import org.school.analysis.service.*;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;
import java.util.List;

import static org.school.analysis.model.ProcessingStatus.*;

@Service
@AllArgsConstructor
@Slf4j
public class ReportProcessorServiceImpl implements ReportProcessorService {

    private final ReportParserService parserService;
    private final StudentResultRepositoryImpl repositoryImpl;
    private final FileOrganizerService fileOrganizerService;

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

            // 2. Парсинг файлов
            List<ParseResult> parseResults = parseReports(foundFiles);
            long successfullyParsed = parseResults.stream()
                    .filter(ParseResult::isSuccess)
                    .count();
            log.info("Успешно распарсено {} файлов", successfullyParsed);

            // 3. Сохранение в БД
            List<ReportFile> savedFiles = saveResultsToDatabase(parseResults);
            summary.setSuccessfullySaved(savedFiles.size());
            log.info("Успешно сохранено {} файлов", savedFiles.size());

            // 4. Перемещение файлов
            List<ReportFile> movedFiles = moveProcessedFiles(savedFiles);
            summary.setSuccessfullyMoved(movedFiles.size());

            // 5. Сбор статистики
            int totalStudents = parseResults.stream()
                    .filter(ParseResult::isSuccess)
                    .mapToInt(ParseResult::getParsedStudents)
                    .sum();
            summary.setTotalStudentsProcessed(totalStudents);
            log.info("Обработано {} студентов", totalStudents);

            // 6. Сбор ошибок
            for (ReportFile file : foundFiles) {
                if (file.getStatus() == ERROR_PARSING ||
                        file.getStatus() == ERROR_SAVING ||
                        file.getStatus() == ERROR_MOVING) {
                    failedFiles.add(file);
                }
            }
            summary.setFailedFiles(failedFiles);

        } catch (Exception e) {
            log.error("Ошибка при обработке отчетов: {}", e.getMessage(), e);
            throw new RuntimeException("Ошибка при обработке отчетов: " + e.getMessage(), e);
        }

        return summary;
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