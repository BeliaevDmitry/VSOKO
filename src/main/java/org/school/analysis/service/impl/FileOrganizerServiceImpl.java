package org.school.analysis.service.impl;

import org.school.analysis.model.ProcessingStatus;
import org.school.analysis.model.ReportFile;
import org.school.analysis.service.FileOrganizerService;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import static org.school.analysis.config.AppConfig.*;

@Service
public class FileOrganizerServiceImpl implements FileOrganizerService {

    private final String reportsBaseFolderTemplate = REPORTS_BASE_FOLDER;

    public FileOrganizerServiceImpl() {

        System.out.println("Используется базовый путь: " + reportsBaseFolderTemplate);
    }

    @Override
    public boolean moveToSubjectFolder(ReportFile reportFile) throws IOException {
        String subject = reportFile.getSubject();
        if (subject == null || subject.isEmpty() || "Неизвестный предмет".equals(subject)) {
            throw new IllegalArgumentException("Не указан предмет для файла");
        }

        String safeSubject = subject.replaceAll("[\\\\/:*?\"<>|]", "_");
        String targetFolderPath = reportsBaseFolderTemplate.replace("{предмет}", safeSubject);

        // Создаем папку если не существует
        Path targetFolder = Paths.get(targetFolderPath);
        if (!Files.exists(targetFolder)) {
            Files.createDirectories(targetFolder);
        }

        // Перемещаем файл
        Path source = reportFile.getFile().toPath();
        Path target = targetFolder.resolve(reportFile.getFile().getName());
        Files.move(source, target, StandardCopyOption.REPLACE_EXISTING);

        return true;
    }

    @Override
    public List<ReportFile> moveFilesToSubjectFolders(List<ReportFile> reportFiles) {
        List<ReportFile> movedFiles = new ArrayList<>();

        for (ReportFile reportFile : reportFiles) {
            try {
                if (moveToSubjectFolder(reportFile)) {
                    movedFiles.add(reportFile);
                }
            } catch (IOException e) {
                // Можно добавить логирование ошибки
                System.err.println("Ошибка при перемещении файла " +
                        reportFile.getFile().getName() + ": " + e.getMessage());
            }
        }

        return movedFiles;
    }

    @Override
    public List<ReportFile> findReportFiles(String folderPath) {
        List<ReportFile> reportFiles = new ArrayList<>();
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            throw new IllegalArgumentException("Папка не существует: " + folderPath);
        }

        // Простой и понятный вариант
        for (File file : folder.listFiles()) {
            if (file.isFile() && file.getName().toLowerCase().endsWith(".xlsx")) {
                ReportFile reportFile = new ReportFile();
                reportFile.setFile(file);
                reportFile.setStatus(ProcessingStatus.PENDING);
                reportFile.setProcessedAt(LocalDateTime.now());

                ReportFile metadata = extractMetadataFromFileName(file);
                if (metadata.getSubject() != null) {
                    reportFile.setSubject(metadata.getSubject());
                }
                if (metadata.getClassName() != null) {
                    reportFile.setClassName(metadata.getClassName());
                }

                reportFiles.add(reportFile);
            }
        }

        return reportFiles;
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