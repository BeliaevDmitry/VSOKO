package org.school.analysis.model;

import lombok.Data;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

@Data
public class ProcessingSummary {
    private int totalFilesFound;
    private int successfullyParsed;
    private int successfullySaved;
    private int successfullyMoved;
    private int generatedReportsCount;
    private List<File> generatedReportFiles = new ArrayList<>();
    private String reportGenerationError;

    public int getFailedToParse() {
        return totalFilesFound - successfullyParsed;
    }

    public int getFailedToSave() {
        return successfullyParsed - successfullySaved;
    }

    public int getFailedToMove() {
        return successfullySaved - successfullyMoved;
    }

    @Override
    public String toString() {
        return String.format(
                "Обработка завершена: Найдено=%d, Распарсено=%d, Сохранено=%d, Перемещено=%d, Отчетов=%d",
                totalFilesFound, successfullyParsed, successfullySaved, successfullyMoved, generatedReportsCount
        );
    }
}