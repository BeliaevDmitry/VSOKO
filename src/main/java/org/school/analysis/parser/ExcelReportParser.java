package org.school.analysis.parser;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.model.ParseResult;
import org.school.analysis.model.ReportFile;
import org.school.analysis.model.StudentResult;
import org.school.analysis.model.TestMetadata;
import org.school.analysis.parser.strategy.MetadataParser;
import org.school.analysis.parser.strategy.StudentDataParser;
import org.school.analysis.util.ExcelParser;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.List;

/**
 * Главный парсер Excel отчетов
 */
public class ExcelReportParser {

    private final MetadataParser metadataParser;
    private final StudentDataParser studentDataParser;

    public ExcelReportParser() {
        this.metadataParser = new MetadataParser();
        this.studentDataParser = new StudentDataParser();
    }

    /**
     * Полный парсинг Excel файла
     */
    public ParseResult parseFile(ReportFile reportFile) throws IOException {
        ParseResult parseResult = new ParseResult();
        parseResult.setReportFile(reportFile);

        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            // 1. Парсинг метаданных
            Sheet infoSheet = workbook.getSheet("Информация");
            TestMetadata metadata = metadataParser.parseMetadata(infoSheet);

            // Если в листе "Информация" нет данных, берем из имени файла
            if (metadata.getSubject() == null || metadata.getClassName() == null) {
                TestMetadata fileMetadata = metadataParser.parseFromFileName(
                        reportFile.getFile().getName()
                );
                metadata.setSubject(fileMetadata.getSubject());
                metadata.setClassName(fileMetadata.getClassName());
            }

            // 2. Парсинг данных учеников
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            if (dataSheet == null) {
                throw new IOException("Лист 'Сбор информации' не найден");
            }

            // 3. Получение максимальных баллов
            var maxScores = studentDataParser.parseMaxScores(dataSheet);
            metadata.setMaxScores(maxScores);
            metadata.setTaskCount(maxScores.size());
            metadata.calculateMaxTotalScore();

            // 4. Парсинг данных учеников
            List<StudentResult> studentResults = studentDataParser.parseStudentData(
                    dataSheet, maxScores, metadata.getSubject(), metadata.getClassName()
            );

            // 5. Установка метаданных для каждого ученика
            for (StudentResult student : studentResults) {
                student.setSubject(metadata.getSubject());
                student.setClassName(metadata.getClassName());
                student.setMaxScores(new java.util.HashMap<>(maxScores));
                student.calculateAll();
            }

            // 6. Формирование результата
            parseResult.setStudentResults(studentResults);
            parseResult.setSuccess(true);
            parseResult.setParsedStudents(studentResults.size());
            parseResult.setMetadata(metadata);
            parseResult.setParsedAt(LocalDateTime.now());

            // 7. Обновление ReportFile
            reportFile.setSubject(metadata.getSubject());
            reportFile.setClassName(metadata.getClassName());
            reportFile.setStudentCount(studentResults.size());

        } catch (Exception e) {
            parseResult.setSuccess(false);
            parseResult.setErrorMessage(e.getMessage());
            throw new IOException("Ошибка парсинга файла: " + reportFile.getFile().getName(), e);
        }

        return parseResult;
    }

    /**
     * Быстрый парсинг только метаданных (без данных учеников)
     */
    public TestMetadata parseMetadataOnly(ReportFile reportFile) throws IOException {
        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet infoSheet = workbook.getSheet("Информация");
            if (infoSheet != null) {
                return metadataParser.parseMetadata(infoSheet);
            }

            return metadataParser.parseFromFileName(reportFile.getFile().getName());

        } catch (Exception e) {
            // Если не удалось прочитать файл, парсим только имя
            return metadataParser.parseFromFileName(reportFile.getFile().getName());
        }
    }

    /**
     * Проверка валидности Excel файла
     */
    public boolean validateExcelFile(ReportFile reportFile) {
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