package org.school.analysis.service.impl;

import org.school.analysis.model.ReportFile;
import org.school.analysis.service.ReportFileFinderService;
import org.springframework.stereotype.Service;

import java.io.File;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import static org.school.analysis.model.ProcessingStatus.PENDING;

@Service
public class ReportFileFinderServiceImpl implements ReportFileFinderService {

    @Override
    public List<ReportFile> findReportFiles(String folderPath) {
        List<ReportFile> reportFiles = new ArrayList<>();
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            throw new IllegalArgumentException("Папка не существует: " + folderPath);
        }

        File[] files = folder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx")
        );

        if (files != null) {
            for (File file : files) {
                ReportFile reportFile = extractMetadataFromFileName(file);
                reportFile.setStatus(PENDING);
                reportFile.setProcessedAt(LocalDateTime.now());
                reportFiles.add(reportFile);
            }
        }

        return reportFiles;
    }

    @Override
    public List<ReportFile> findFilesByPattern(String folderPath, String pattern) {
        List<ReportFile> allFiles = findReportFiles(folderPath);
        List<ReportFile> filteredFiles = new ArrayList<>();

        for (ReportFile reportFile : allFiles) {
            if (reportFile.getFile().getName().contains(pattern)) {
                filteredFiles.add(reportFile);
            }
        }

        return filteredFiles;
    }

    @Override
    public ReportFile extractMetadataFromFileName(File file) {
        ReportFile reportFile = new ReportFile();
        reportFile.setFile(file);

        // Пример: "Сбор_данных_10А_История.xlsx"
        String[] parts = file.getName().replace(".xlsx", "").split("_");

        // Ищем класс (формат: "10А", "6", "11Б")
        for (String part : parts) {
            if (part.matches("\\d+[А-Яа-я]?")) {
                reportFile.setClassName(part);
                break;
            }
        }

        // Ищем предмет (последняя часть, не класс, не "Сбор", не "данных")
        for (int i = parts.length - 1; i >= 0; i--) {
            String part = parts[i];
            if (!part.matches("\\d+[А-Яа-я]?") &&
                    !part.equalsIgnoreCase("Сбор") &&
                    !part.equalsIgnoreCase("данных") &&
                    !part.equalsIgnoreCase("класс")) {
                reportFile.setSubject(part);
                break;
            }
        }

        // Если не нашли, устанавливаем по умолчанию
        if (reportFile.getSubject() == null) {
            reportFile.setSubject("Неизвестный предмет");
        }
        if (reportFile.getClassName() == null) {
            reportFile.setClassName("Неизвестный класс");
        }

        return reportFile;
    }
}