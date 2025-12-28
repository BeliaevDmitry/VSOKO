package org.school.analysis.service.impl;

import org.school.analysis.model.ReportFile;
import org.school.analysis.service.FileOrganizerService;

import java.io.IOException;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.List;

public class FileOrganizerServiceImpl implements FileOrganizerService {

    private final String reportsBaseFolderTemplate;

    public FileOrganizerServiceImpl(String reportsBaseFolderTemplate) {
        this.reportsBaseFolderTemplate = reportsBaseFolderTemplate;
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
    public String createSubjectFolderStructure(String subject) throws IOException {
        String safeSubject = subject.replaceAll("[\\\\/:*?\"<>|]", "_");
        String targetFolderPath = reportsBaseFolderTemplate.replace("{предмет}", safeSubject);

        Path targetFolder = Paths.get(targetFolderPath);
        if (!Files.exists(targetFolder)) {
            Files.createDirectories(targetFolder);
        }

        return targetFolder.toString();
    }
}