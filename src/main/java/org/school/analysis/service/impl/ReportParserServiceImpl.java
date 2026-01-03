package org.school.analysis.service.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.exception.ValidationException;
import org.school.analysis.model.*;
import org.school.analysis.parser.strategy.MetadataParser;
import org.school.analysis.parser.strategy.StudentDataParser;
import org.school.analysis.service.ReportParserService;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Главный парсер Excel отчетов
 */
@Service
public class ReportParserServiceImpl implements ReportParserService {

    private final MetadataParser metadataParser;
    private final StudentDataParser studentDataParser;

    public ReportParserServiceImpl(MetadataParser metadataParser, StudentDataParser studentDataParser) {
        this.metadataParser = metadataParser;
        this.studentDataParser = studentDataParser;
    }

    /**
     * Парсинг списка файлов
     */
    @Override
    public List<ParseResult> parseFiles(List<ReportFile> reportFiles) {
        return reportFiles.stream()
                .map(this::parseFile)
                .collect(Collectors.toList());
    }

    /**
     * Полный парсинг Excel файла
     */
    @Override
    public ParseResult parseFile(ReportFile reportFile) {
        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            // 1. Парсинг метаданных
            if (!validateExcelFile(reportFile)) {
                return ParseResult.error(reportFile, "неправильная структура отчёта");
            }
            Sheet infoSheet = workbook.getSheet("Информация");
            TestMetadata metadata = metadataParser.parseMetadata(infoSheet);

            // 2. Парсинг данных учеников
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            if (dataSheet == null) {
                return ParseResult.error(reportFile, "Лист 'Сбор информации' не найден");
            }

            // 3. Получение максимальных баллов
            var maxScores = studentDataParser.parseMaxScores(dataSheet);

            // 4. Парсинг данных учеников
            List<StudentResult> studentResults = studentDataParser.parseStudentData(
                    dataSheet, maxScores, metadata.getSubject(), metadata.getClassName());

            // 5. ПОЛНОЕ обновление ReportFile из TestMetadata
            reportFile.setSubject(metadata.getSubject());
            reportFile.setClassName(metadata.getClassName());
            reportFile.setTestDate(metadata.getTestDate());
            reportFile.setTeacher(metadata.getTeacher());
            reportFile.setSchool(metadata.getSchool() != null ? metadata.getSchool() : "ГБОУ №7");
            reportFile.setTestType(metadata.getTestType());
            reportFile.setComment(metadata.getComment());

            // Параметры теста
            reportFile.setTaskCount(maxScores.size());
            reportFile.setMaxScores(maxScores);
            metadata.setMaxScores(maxScores);  // ДОБАВИТЬ эту строку!
            metadata.calculateMaxTotalScore();  // обновить metadata
            reportFile.setMaxTotalScore(metadata.getMaxTotalScore());  // взять из metadata
            reportFile.setStudentCount(studentResults.size());// Статистика

            // 6. Установка метаданных для каждого ученика
            for (StudentResult student : studentResults) {
                student.setSubject(reportFile.getSubject());
                student.setClassName(reportFile.getClassName());
                student.setTestDate(reportFile.getTestDate());
                student.setTestType(reportFile.getTestType());
            }

            // 7. Формирование успешного результата
            return ParseResult.success(reportFile, studentResults);

        } catch (ValidationException e) {
            return ParseResult.error(reportFile,
                    "Ошибка валидации данных. Проверьте файл:\n" + e.getMessage());

        } catch (Exception e) {
            return ParseResult.error(reportFile,
                    "Ошибка парсинга файла " + reportFile.getFile().getName() + ": " + e.getMessage());
        }
    }

    /**
     * Проверка валидности Excel файла
     */
    private boolean validateExcelFile(ReportFile reportFile) {
        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            // Проверяем наличие необходимых листов
            boolean hasInfoSheet = workbook.getSheet("Информация") != null;
            boolean hasDataSheet = workbook.getSheet("Сбор информации") != null;

            return hasInfoSheet && hasDataSheet;

        } catch (Exception e) {
            return false;
        }
    }
}