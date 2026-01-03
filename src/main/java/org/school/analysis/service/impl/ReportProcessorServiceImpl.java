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

    // Добавьте это для контроля размера пакета
    private static final int BATCH_SIZE = 100;

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
            for (int i = 0; i < foundFiles.size(); i += BATCH_SIZE) {
                int end = Math.min(i + BATCH_SIZE, foundFiles.size());
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

            // 3. Итоговая статистика
            log.info("ИТОГО: Найдено={}, Распарсено={}, Сохранено={}, Перемещено={}",
                    summary.getTotalFilesFound(),
                    summary.getSuccessfullyParsed(),
                    summary.getSuccessfullySaved(),
                    summary.getSuccessfullyMoved());

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