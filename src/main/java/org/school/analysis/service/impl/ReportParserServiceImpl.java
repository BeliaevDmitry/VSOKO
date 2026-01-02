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
import java.io.IOException;
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
            Sheet infoSheet = workbook.getSheet("Информация");
            TestMetadata metadata = metadataParser.parseMetadata(infoSheet);

            // 2. Парсинг данных учеников
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            if (dataSheet == null) {
                return ParseResult.error(reportFile, "Лист 'Сбор информации' не найден");
            }

            // 3. Получение максимальных баллов
            var maxScores = studentDataParser.parseMaxScores(dataSheet);
            metadata.setMaxScores(maxScores);
            metadata.setTaskCount(maxScores.size());
            metadata.calculateMaxTotalScore();

            // 4. Парсинг данных учеников (может бросить ValidationException)
            List<StudentResult> studentResults = studentDataParser.parseStudentData(
                    dataSheet, maxScores, metadata.getSubject(), metadata.getClassName());

            // 5. Установка метаданных для каждого ученика
            for (StudentResult student : studentResults) {
                student.setSubject(metadata.getSubject());
                student.setClassName(metadata.getClassName());
                student.setTestDate(metadata.getTestDate());
            }

            // 6. Обновление ReportFile
            reportFile.setSubject(metadata.getSubject());
            reportFile.setClassName(metadata.getClassName());
            reportFile.setStudentCount(studentResults.size());

            // 7. Формирование успешного результата
            return ParseResult.success(reportFile, studentResults);

        } catch (ValidationException e) {
            // Специальная обработка ошибок валидации
            return ParseResult.error(reportFile,
                    "Ошибка валидации данных. Проверьте файл:\n" + e.getMessage());

        } catch (Exception e) {
            return ParseResult.error(reportFile,
                    "Ошибка парсинга файла " + reportFile.getFile().getName() + ": " + e.getMessage());
        }
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