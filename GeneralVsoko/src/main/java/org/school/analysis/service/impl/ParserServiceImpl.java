package org.school.analysis.service.impl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.exception.ValidationException;
import org.school.analysis.model.*;
import org.school.analysis.parser.strategy.MetadataParser;
import org.school.analysis.parser.strategy.StudentDataParser;
import org.school.analysis.service.ParserService;
import org.school.analysis.util.JsonScoreUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Главный парсер Excel отчетов
 */
@Service
public class ParserServiceImpl implements ParserService {

    private static final Logger log = LoggerFactory.getLogger(ParserServiceImpl.class);

    private final MetadataParser metadataParser;
    private final StudentDataParser studentDataParser;

    public ParserServiceImpl(MetadataParser metadataParser, StudentDataParser studentDataParser) {
        this.metadataParser = metadataParser;
        this.studentDataParser = studentDataParser;
    }

    /**
     * Парсинг списка файлов
     */
    @Override
    public List<ParseResult> parseFiles(List<ReportFile> reportFiles) {
        log.info("Начало парсинга {} файлов", reportFiles.size());
        return reportFiles.stream()
                .map(this::parseFile)
                .collect(Collectors.toList());
    }

    /**
     * Полный парсинг Excel файла
     */
    @Override
    public ParseResult parseFile(ReportFile reportFile) {
        log.info("Начинаем парсинг файла: {}", reportFile.getFile().getName());
        log.debug("Полный путь к файлу: {}", reportFile.getFile().getAbsolutePath());

        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            log.debug("Файл успешно открыт, количество листов: {}", workbook.getNumberOfSheets());

            // Выводим названия всех листов для отладки
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                log.debug("Лист {}: {}", i, workbook.getSheetName(i));
            }

            // 1. Проверяем файл на структуру и парсим лист Информация
            log.debug("Проверка структуры файла...");
            if (!validateExcelFile(reportFile)) {
                log.error("Файл {} не прошел валидацию - неправильная структура", reportFile.getFile().getName());
                return ParseResult.error(reportFile, "неправильная структура отчёта");
            }

            Sheet infoSheet = workbook.getSheet("Информация");
            log.debug("Лист 'Информация' найден, строк: {}", infoSheet.getPhysicalNumberOfRows());

            TestMetadata metadata = metadataParser.parseMetadata(infoSheet);
            var maxScores = metadata.getMaxScores();
            log.debug("Парсинг максимальных баллов...");
            log.info("Метаданные распарсены: предмет={}, класс={}, дата={}, учитель={}",
                    metadata.getSubject(), metadata.getClassName(),
                    metadata.getTestDate(), metadata.getTeacher());

            // 2. Парсинг данных учеников
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            log.debug("Лист 'Сбор информации' найден, строк: {}", dataSheet.getPhysicalNumberOfRows());

            log.debug("Парсинг данных учеников...");
            List<StudentResult> studentResults = studentDataParser.parseStudentData(
                    dataSheet, maxScores, metadata.getSubject(), metadata.getClassName());

            // 3. ПОЛНОЕ обновление ReportFile из TestMetadata
            log.debug("Обновление информации о файле...");
            reportFile.setSubject(metadata.getSubject());
            reportFile.setClassName(metadata.getClassName());
            reportFile.setTestDate(metadata.getTestDate());
            reportFile.setTeacher(metadata.getTeacher());
            reportFile.setSchool(metadata.getSchool() != null ? metadata.getSchool() : "ГБОУ №7");
            reportFile.setTestType(metadata.getTestType());
            reportFile.setComment(metadata.getComment());
            reportFile.setACADEMIC_YEAR(metadata.getACADEMIC_YEAR() != null ? metadata.getACADEMIC_YEAR() : "2025-2026");
            reportFile.setTaskCount(maxScores.size());
            reportFile.setMaxScores(maxScores);

            reportFile.setStudentCount(studentResults.size());

            log.debug("Информация о файле обновлена: заданий={}, учеников={}",
                    reportFile.getTaskCount(), reportFile.getStudentCount());

            // 4. Установка метаданных для каждого ученика
            log.debug("Установка метаданных для учеников...");
            for (StudentResult student : studentResults) {
                student.setSubject(reportFile.getSubject());
                student.setClassName(reportFile.getClassName());
                student.setTestDate(reportFile.getTestDate());
                student.setTestType(reportFile.getTestType());
                student.setSchool(reportFile.getSchool());
                student.setACADEMIC_YEAR(reportFile.getACADEMIC_YEAR());

                // Вычисляем totalScore для каждого студента
                if (student.getTaskScores() != null) {
                    int totalScore = JsonScoreUtils.calculateTotalScore(student.getTaskScores());
                    student.setTotalScore(totalScore);

                    // Вычисляем процент выполнения
                    if (!maxScores.isEmpty() && reportFile.getMaxTotalScore() > 0) {
                        double percentage = (totalScore * 100.0) / reportFile.getMaxTotalScore();
                        student.setPercentageScore(Math.round(percentage * 100.0) / 100.0);
                    }
                }
            }

            // 5. Формирование успешного результата
            log.info("Файл {} успешно обработан: {} учеников, {} заданий",
                    reportFile.getFile().getName(), studentResults.size(), maxScores.size());
            return ParseResult.success(reportFile, studentResults);

        } catch (ValidationException e) {
            log.error("Ошибка валидации в файле {}: {}",
                    reportFile.getFile().getName(), e.getMessage(), e);
            return ParseResult.error(reportFile,
                    "Ошибка валидации данных. Проверьте файл:\n" + e.getMessage());

        } catch (Exception e) {
            log.error("Критическая ошибка парсинга файла {}: {}",
                    reportFile.getFile().getName(), e.getMessage(), e);
            return ParseResult.error(reportFile,
                    "Ошибка парсинга файла " + reportFile.getFile().getName() + ": " + e.getMessage());
        }
    }

    /**
     * Проверка валидности Excel файла
     */
    private boolean validateExcelFile(ReportFile reportFile) {
        log.debug("Валидация структуры файла: {}", reportFile.getFile().getName());

        try (FileInputStream file = new FileInputStream(reportFile.getFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            // Проверяем наличие необходимых листов
            boolean hasInfoSheet = workbook.getSheet("Информация") != null;
            boolean hasDataSheet = workbook.getSheet("Сбор информации") != null;

            log.debug("Результаты валидации: hasInfoSheet={}, hasDataSheet={}",
                    hasInfoSheet, hasDataSheet);

            return hasInfoSheet && hasDataSheet;

        } catch (Exception e) {
            log.error("Ошибка при валидации файла {}: {}",
                    reportFile.getFile().getName(), e.getMessage());
            return false;
        }
    }

    /**
     * Вспомогательный метод для логирования первых строк листа
     */
    private void logFirstRows(Sheet sheet, int count) {
        log.debug("=== Первые {} строк листа '{}' ===", count, sheet.getSheetName());

        for (int i = 0; i <= Math.min(count, sheet.getLastRowNum()); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                StringBuilder rowData = new StringBuilder();
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        rowData.append(getCellValue(cell)).append(" | ");
                    }
                }
                log.debug("Строка {}: {}", i, rowData.toString());
            }
        }
        log.debug("=== Конец отладки строк ===");
    }

    /**
     * Получение значения ячейки в виде строки
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "[ПУСТО]";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "[ПУСТО]";
            default:
                return "[НЕИЗВЕСТНО]";
        }
    }
}