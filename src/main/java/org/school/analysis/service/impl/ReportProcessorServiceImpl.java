package org.school.analysis.service.impl;

import lombok.AllArgsConstructor;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ReportFile;
import org.school.analysis.repository.StudentResultRepository;
import org.school.analysis.service.*;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

import static org.school.analysis.model.ProcessingStatus.*;

@Service
@AllArgsConstructor
public class ReportProcessorServiceImpl implements ReportProcessorService {

    private final ReportFileFinderService fileFinderService;
    private final ReportParserService parserService;
    private final StudentResultRepository repository;
    private final FileOrganizerService fileOrganizerService;

    @Override
    public ProcessingSummary processAll(String folderPath) {
        ProcessingSummary summary = new ProcessingSummary();
        List<ReportFile> failedFiles = new ArrayList<>();

        try {
            // 1. Найти файлы
            List<ReportFile> foundFiles = findReports(folderPath);
            summary.setTotalFilesFound(foundFiles.size());

            // 2. Парсинг файлов
            List<ParseResult> parseResults = parseReports(foundFiles);

            // 3. Сохранение в БД
            List<ReportFile> savedFiles = saveResultsToDatabase(parseResults);
            int successfullyParsedCount = (int) parseResults.stream()
                    .filter(ParseResult::isSuccess)
                    .count();
            summary.setSuccessfullySaved(savedFiles.size());

            // 4. Перемещение файлов
            List<ReportFile> movedFiles = moveProcessedFiles(savedFiles);
            summary.setSuccessfullyMoved(movedFiles.size());

            // 5. Сбор статистики
            summary.setTotalStudentsProcessed(
                    parseResults.stream()
                            .filter(ParseResult::isSuccess)
                            .mapToInt(ParseResult::getParsedStudents)
                            .sum()
            );

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
            throw new RuntimeException("Ошибка при обработке отчетов: " + e.getMessage(), e);
        }

        return summary;
    }

    @Override
    public List<ReportFile> findReports(String folderPath) {
        return fileFinderService.findReportFiles(folderPath);
    }

    @Override
    public List<ParseResult> parseReports(List<ReportFile> reportFiles) {
        return parserService.parseFiles(reportFiles);
    }

    @Override
    public List<ReportFile> saveResultsToDatabase(List<ParseResult> parseResults) {
        List<ReportFile> successfullySaved = new ArrayList<>();

        for (ParseResult parseResult : parseResults) {
            if (parseResult.isSuccess()) {
                int savedCount = repository.saveAll(parseResult.getStudentResults());
                if (savedCount > 0) {
                    parseResult.getReportFile().setStatus(SAVED);
                    successfullySaved.add(parseResult.getReportFile());
                } else {
                    parseResult.getReportFile().setStatus(ERROR_SAVING);
                    parseResult.getReportFile().setErrorMessage("Не удалось сохранить в БД");
                }
            }
        }

        return successfullySaved;
    }

    @Override
    public List<ReportFile> moveProcessedFiles(List<ReportFile> successfullyProcessedFiles) {
        return fileOrganizerService.moveFilesToSubjectFolders(successfullyProcessedFiles);
    }
}