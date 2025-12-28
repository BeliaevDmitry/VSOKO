package org.school.analysis.service.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.*;
import org.school.analysis.service.ReportParserService;
import org.school.analysis.util.ExcelParser;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.school.analysis.model.ProcessingStatus.*;

public class ReportParserServiceImpl implements ReportParserService {

    private final ParsingStatistics statistics = new ParsingStatistics();

    @Override
    public ParseResult parseFile(ReportFile reportFile) {
        ParseResult parseResult = new ParseResult();
        parseResult.setReportFile(reportFile);
        reportFile.setStatus(PARSING);

        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            // 1. Получить метаданные
            Sheet infoSheet = workbook.getSheet("Информация");
            if (infoSheet != null) {
                String subject = ExcelParser.getCellValueAsString(infoSheet, 2, 1, "Неизвестный предмет");
                String className = ExcelParser.getCellValueAsString(infoSheet, 3, 1, "Неизвестный класс");
                reportFile.setSubject(subject);
                reportFile.setClassName(className);
            }

            // 2. Парсинг данных учеников
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            if (dataSheet == null) {
                throw new Exception("Лист 'Сбор информации' не найден");
            }

            // 3. Чтение максимальных баллов
            Map<Integer, Integer> maxScores = readMaxScores(dataSheet);

            // 4. Парсинг учеников
            List<StudentResult> studentResults = parseStudentData(dataSheet, maxScores);

            // 5. Установка результатов
            parseResult.setStudentResults(studentResults);
            parseResult.setSuccess(true);
            parseResult.setParsedStudents(studentResults.size());
            reportFile.setStatus(PARSED);
            reportFile.setStudentCount(studentResults.size());

            statistics.successfullyParsed++;
            statistics.totalStudents += studentResults.size();

        } catch (Exception e) {
            parseResult.setSuccess(false);
            parseResult.setErrorMessage(e.getMessage());
            reportFile.setStatus(ERROR_PARSING);
            reportFile.setErrorMessage(e.getMessage());
            statistics.failedParsed++;
        }

        statistics.totalFiles++;
        return parseResult;
    }

    @Override
    public List<ParseResult> parseFiles(List<ReportFile> reportFiles) {
        List<ParseResult> results = new ArrayList<>();
        for (ReportFile reportFile : reportFiles) {
            results.add(parseFile(reportFile));
        }
        return results;
    }

    @Override
    public ParsingStatistics getStatistics() {
        return statistics;
    }

    private Map<Integer, Integer> readMaxScores(Sheet sheet) {
        Map<Integer, Integer> maxScores = new HashMap<>();
        Row maxScoresRow = sheet.getRow(2); // 3-я строка с макс. баллами

        if (maxScoresRow != null) {
            int taskNumber = 1;
            for (int col = 4; col < 100; col++) {
                Cell cell = maxScoresRow.getCell(col);
                if (cell == null) break;

                Integer maxScore = ExcelParser.getCellValueAsInteger(cell);
                if (maxScore != null) {
                    maxScores.put(taskNumber, maxScore);
                    taskNumber++;
                } else {
                    break;
                }
            }
        }

        return maxScores;
    }

    private List<StudentResult> parseStudentData(Sheet sheet, Map<Integer, Integer> maxScores) {
        List<StudentResult> results = new ArrayList<>();
        int firstStudentRow = 3;
        int maxStudents = 34;

        for (int rowIdx = firstStudentRow; rowIdx < firstStudentRow + maxStudents; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) break;

            String fio = ExcelParser.getCellValueAsString(row.getCell(1));
            if (fio == null || fio.trim().isEmpty()) continue;

            String presence = ExcelParser.getCellValueAsString(row.getCell(2));
            if ("Не был".equalsIgnoreCase(presence)) continue;

            StudentResult result = new StudentResult();
            result.setFio(fio.trim());
            result.setPresence(presence);
            result.setVariant(ExcelParser.getCellValueAsString(row.getCell(3)));

            // Парсинг баллов за задания
            Map<Integer, Integer> taskScores = new HashMap<>();
            int totalScore = 0;

            for (Map.Entry<Integer, Integer> entry : maxScores.entrySet()) {
                int taskNum = entry.getKey();
                int columnIndex = 3 + taskNum;
                Cell scoreCell = row.getCell(columnIndex);

                Integer score = ExcelParser.getCellValueAsInteger(scoreCell);
                if (score == null) score = 0;

                taskScores.put(taskNum, score);
                totalScore += score;
            }

            result.setTaskScores(taskScores);
            result.setTotalScore(totalScore);
            result.setMaxScores(new HashMap<>(maxScores));
            result.calculateAll();

            results.add(result);
        }

        return results;
    }
}