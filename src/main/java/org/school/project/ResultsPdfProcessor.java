package org.school.project;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.*;

public class ResultsPdfProcessor {

    private static boolean DEBUG = true; // Установите true для отладки

    public static void main(String[] args) {
        // Укажите путь к папке с PDF файлами результатов
        String folderPath = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\МЦКО\\на обработку\\результаты";
        // Укажите путь для сохранения Excel файла
        String outputExcelPath = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\МЦКО\\код_результат.xlsx";

        try {
            List<StudentResultData> allResults = new ArrayList<>();

            System.out.println("Начало обработки PDF файлов из папки: " + folderPath);
            System.out.println("==========================================");

            // Получаем список всех PDF файлов в папке
            List<Path> pdfFiles = new ArrayList<>();
            try (var paths = Files.walk(Paths.get(folderPath))) {
                paths.filter(Files::isRegularFile)
                        .filter(path -> path.toString().toLowerCase().endsWith(".pdf"))
                        .forEach(pdfFiles::add);
            }

            System.out.println("Найдено PDF файлов: " + pdfFiles.size());

            // Обрабатываем каждый PDF файл
            for (Path pdfPath : pdfFiles) {
                try {
                    System.out.println("\nОбработка файла: " + pdfPath.getFileName());
                    processResultPdfFile(pdfPath.toFile(), allResults);
                } catch (Exception e) {
                    System.err.println("Ошибка при обработке файла: " + pdfPath);
                    e.printStackTrace();
                }
            }

            System.out.println("\n==========================================");
            System.out.println("Всего извлечено записей: " + allResults.size());

            // Создаем Excel файл
            createResultsExcelFile(allResults, outputExcelPath);
            System.out.println("Обработка завершена. Результат сохранен в: " + outputExcelPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void processResultPdfFile(File pdfFile, List<StudentResultData> allResults) throws IOException {
        String fileName = pdfFile.getName();

        try (PDDocument document = PDDocument.load(pdfFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);

            if (DEBUG) {
                System.out.println("  Первые 500 символов текста PDF:");
                System.out.println(text.substring(0, Math.min(500, text.length())));
                System.out.println("  ---");
            }

            // Извлекаем общую информацию из заголовка
            Map<String, String> headerInfo = extractHeaderInfo(text);
            String date = headerInfo.get("date");
            String subject = headerInfo.get("subject");
            String className = headerInfo.get("className");

            if (DEBUG) {
                System.out.println("  Дата: " + date);
                System.out.println("  Предмет: " + subject);
                System.out.println("  Класс: " + className);
            }

            // Извлекаем данные студентов
            List<StudentResultData> studentResults = extractStudentResults(text, className, subject, date);

            System.out.println("  Найдено записей студентов: " + studentResults.size());
            allResults.addAll(studentResults);

            if (DEBUG && studentResults.size() > 0) {
                System.out.println("  Первые 3 записи:");
                for (int i = 0; i < Math.min(3, studentResults.size()); i++) {
                    StudentResultData s = studentResults.get(i);
                    System.out.println("    Код: " + s.getCode() +
                            " | Уровень: " + s.getMasteryLevel() +
                            " | Раздел1: " + s.getSection1Percent() +
                            " | Раздел2: " + s.getSection2Percent() +
                            " | Раздел3: " + s.getSection3Percent());
                }
            }
        }
    }

    private static Map<String, String> extractHeaderInfo(String text) {
        Map<String, String> info = new HashMap<>();
        info.put("date", "");
        info.put("subject", "");
        info.put("className", "");

        String[] lines = text.split("\n");

        for (String line : lines) {
            line = line.trim();

            // Ищем дату в формате "Дата: 11-12 ноября 2025г."
            if (line.contains("Дата:")) {
                Pattern datePattern = Pattern.compile("Дата:\\s*(.+?)\\s*(?=Функциональная|Округ:|$)");
                Matcher matcher = datePattern.matcher(line);
                if (matcher.find()) {
                    info.put("date", matcher.group(1).trim());
                }

                // Ищем предмет
                if (line.contains("Функциональная грамотность")) {
                    info.put("subject", "Функциональная грамотность");
                } else {
                    // Ищем другие предметы
                    Pattern subjectPattern = Pattern.compile("Дата:.+?\\s+(.+?)\\s+Округ:");
                    Matcher subjectMatcher = subjectPattern.matcher(line);
                    if (subjectMatcher.find()) {
                        info.put("subject", subjectMatcher.group(1).trim());
                    }
                }

                // Ищем класс
                Pattern classPattern = Pattern.compile("Класс:\\s*(\\d+[А-Яа-яЁё])");
                Matcher classMatcher = classPattern.matcher(line);
                if (classMatcher.find()) {
                    info.put("className", classMatcher.group(1).trim());
                }
            }
        }

        // Если не нашли в одной строке, ищем отдельно
        if (info.get("className").isEmpty()) {
            for (String line : lines) {
                Pattern classPattern = Pattern.compile("Класс:\\s*(\\d+[А-Яа-яЁё])");
                Matcher classMatcher = classPattern.matcher(line);
                if (classMatcher.find()) {
                    info.put("className", classMatcher.group(1).trim());
                    break;
                }
            }
        }

        return info;
    }

    private static List<StudentResultData> extractStudentResults(String text,
                                                                 String className,
                                                                 String subject,
                                                                 String date) {
        List<StudentResultData> results = new ArrayList<>();
        String[] lines = text.split("\n");

        boolean inResultsSection = false;
        int lineNumber = 0;

        for (int i = 0; i < lines.length; i++) {
            String line = lines[i].trim();

            // Находим начало таблицы с результатами
            if (line.matches("\\d+\\s+\\d{4}-\\d{4}\\s+\\d+.*") ||
                    (line.matches(".*\\d{4}-\\d{4}.*") && line.contains("+"))) {
                inResultsSection = true;
            }

            // Ищем строку с заголовками таблицы
            if (line.contains("Фамилия, имя") && line.contains("Код")) {
                inResultsSection = false; // Это заголовок, еще не данные
            }

            if (inResultsSection && !line.isEmpty()) {
                // Пытаемся распарсить строку с данными студента
                StudentResultData result = parseStudentResultLine(line, className, subject, date);
                if (result != null) {
                    results.add(result);
                }
            }

            // Альтернативный поиск: строка начинается с номера и содержит код
            if (!inResultsSection && line.matches("^\\d+\\s+\\d{4}-\\d{4}.*")) {
                StudentResultData result = parseStudentResultLine(line, className, subject, date);
                if (result != null) {
                    results.add(result);
                }
            }
        }

        // Если не нашли данные обычным способом, пытаемся альтернативным парсингом
        if (results.isEmpty()) {
            results = alternativeParse(text, className, subject, date);
        }

        return results;
    }

    private static StudentResultData parseStudentResultLine(String line,
                                                            String className,
                                                            String subject,
                                                            String date) {
        try {
            // Удаляем лишние пробелы
            line = line.replaceAll("\\s+", " ").trim();

            // Паттерн для поиска кода участника
            Pattern codePattern = Pattern.compile("(\\d{4}-\\d{4})");
            Matcher codeMatcher = codePattern.matcher(line);

            if (!codeMatcher.find()) {
                return null;
            }

            String code = codeMatcher.group(1);

            // Ищем проценты выполнения по разделам
            // Паттерн для поиска процентов (число с % или без)
            Pattern percentPattern = Pattern.compile("(\\d{1,3})%");
            Matcher percentMatcher = percentPattern.matcher(line);

            List<String> percents = new ArrayList<>();
            while (percentMatcher.find()) {
                percents.add(percentMatcher.group(1));
            }

            // Ищем уровень овладения УУД
            String masteryLevel = "";
            String[] levelKeywords = {"Повышенный", "Базовый", "Ниже базового", "Высокий", "Средний", "Низкий"};
            for (String level : levelKeywords) {
                if (line.contains(level)) {
                    masteryLevel = level;
                    break;
                }
            }

            // Если уровень не найден, ищем в конце строки
            if (masteryLevel.isEmpty()) {
                String[] words = line.split("\\s+");
                if (words.length > 0) {
                    String lastWord = words[words.length - 1];
                    for (String level : levelKeywords) {
                        if (lastWord.contains(level)) {
                            masteryLevel = level;
                            break;
                        }
                    }
                }
            }

            // Ищем проценты по разделам
            String section1Percent = "";
            String section2Percent = "";
            String section3Percent = "";

            // Ищем паттерн типа "71% 100% 88% 33% 75% 71% 100%"
            // Или "Раздел 1 Раздел 2 Раздел 3"
            // Сначала попробуем найти все числа с процентами
            if (percents.size() >= 7) {
                // Предполагаем, что последние 3 процента - это разделы
                section1Percent = percents.size() >= 3 ? percents.get(percents.size() - 3) + "%" : "";
                section2Percent = percents.size() >= 2 ? percents.get(percents.size() - 2) + "%" : "";
                section3Percent = percents.size() >= 1 ? percents.get(percents.size() - 1) + "%" : "";
            }

            // Альтернативный поиск: ищем конкретно разделы
            if (section1Percent.isEmpty()) {
                // Ищем паттерн с разделителями
                Pattern sectionPattern = Pattern.compile("(\\d{1,3})%\\s*(\\d{1,3})%\\s*(\\d{1,3})%\\s*[А-Яа-яЁё]+$");
                Matcher sectionMatcher = sectionPattern.matcher(line);
                if (sectionMatcher.find()) {
                    section1Percent = sectionMatcher.group(1) + "%";
                    section2Percent = sectionMatcher.group(2) + "%";
                    section3Percent = sectionMatcher.group(3) + "%";
                }
            }

            // Создаем объект с результатами
            StudentResultData result = new StudentResultData();
            result.setCode(code);
            result.setClassName(className);
            result.setSubject(subject);
            result.setDate(date);
            result.setMasteryLevel(masteryLevel);
            result.setSection1Percent(section1Percent.isEmpty() ? "0%" : section1Percent);
            result.setSection2Percent(section2Percent.isEmpty() ? "0%" : section2Percent);
            result.setSection3Percent(section3Percent.isEmpty() ? "0%" : section3Percent);

            return result;

        } catch (Exception e) {
            if (DEBUG) {
                System.err.println("Ошибка парсинга строки: " + line);
                e.printStackTrace();
            }
            return null;
        }
    }

    private static List<StudentResultData> alternativeParse(String text,
                                                            String className,
                                                            String subject,
                                                            String date) {
        List<StudentResultData> results = new ArrayList<>();

        // Разбиваем текст на строки
        String[] lines = text.split("\n");

        for (String line : lines) {
            line = line.trim();

            // Ищем строки, содержащие код участника
            if (line.matches(".*\\d{4}-\\d{4}.*")) {
                // Пропускаем заголовки
                if (line.contains("Код участника") ||
                        line.contains("Фамилия, имя") ||
                        line.contains("№ уч.")) {
                    continue;
                }

                // Пытаемся разобрать строку как данные студента
                String[] parts = line.split("\\s+");
                if (parts.length >= 3) {
                    // Ищем код в формате 9116-0195
                    String code = "";
                    for (String part : parts) {
                        if (part.matches("\\d{4}-\\d{4}")) {
                            code = part;
                            break;
                        }
                    }

                    if (!code.isEmpty()) {
                        // Ищем уровень овладения
                        String masteryLevel = "";
                        String[] levelKeywords = {"Повышенный", "Базовый", "Ниже базового", "Высокий", "Средний", "Низкий"};
                        for (String level : levelKeywords) {
                            if (line.contains(level)) {
                                masteryLevel = level;
                                break;
                            }
                        }

                        // Ищем проценты
                        String section1Percent = "0%";
                        String section2Percent = "0%";
                        String section3Percent = "0%";

                        // Ищем все проценты в строке
                        Pattern percentPattern = Pattern.compile("(\\d{1,3})%");
                        Matcher percentMatcher = percentPattern.matcher(line);
                        List<String> percents = new ArrayList<>();

                        while (percentMatcher.find()) {
                            percents.add(percentMatcher.group(1) + "%");
                        }

                        if (percents.size() >= 3) {
                            section1Percent = percents.get(percents.size() - 3);
                            section2Percent = percents.get(percents.size() - 2);
                            section3Percent = percents.get(percents.size() - 1);
                        }

                        // Создаем объект
                        StudentResultData result = new StudentResultData();
                        result.setCode(code);
                        result.setClassName(className);
                        result.setSubject(subject);
                        result.setDate(date);
                        result.setMasteryLevel(masteryLevel.isEmpty() ? "Не определен" : masteryLevel);
                        result.setSection1Percent(section1Percent);
                        result.setSection2Percent(section2Percent);
                        result.setSection3Percent(section3Percent);

                        results.add(result);
                    }
                }
            }
        }

        return results;
    }

    private static void createResultsExcelFile(List<StudentResultData> results, String outputPath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Результаты участников");

        // Стиль для заголовков
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short)12);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderBottom(BorderStyle.THIN);

        // Создаем заголовки
        Row headerRow = sheet.createRow(0);
        String[] headers = {
                "№",
                "Код участника",
                "Класс",
                "Предмет",
                "Дата проведения",
                "% выполнения заданий по разделу 1",
                "% выполнения заданий по разделу 2",
                "% выполнения заданий по разделу 3",
                "Уровень овладения УУД"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Заполняем данные
        int rowNum = 1;
        for (int i = 0; i < results.size(); i++) {
            StudentResultData result = results.get(i);
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(i + 1); // Номер по порядку
            row.createCell(1).setCellValue(result.getCode());
            row.createCell(2).setCellValue(result.getClassName());
            row.createCell(3).setCellValue(result.getSubject());
            row.createCell(4).setCellValue(result.getDate());
            row.createCell(5).setCellValue(result.getSection1Percent());
            row.createCell(6).setCellValue(result.getSection2Percent());
            row.createCell(7).setCellValue(result.getSection3Percent());
            row.createCell(8).setCellValue(result.getMasteryLevel());
        }

        // Автоподбор ширины колонок
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
            // Устанавливаем минимальную ширину
            if (sheet.getColumnWidth(i) < 3000) {
                sheet.setColumnWidth(i, 3000);
            }
        }

        // Замораживаем строку с заголовками
        sheet.createFreezePane(0, 1);

        // Сохраняем файл
        try (FileOutputStream outputStream = new FileOutputStream(outputPath)) {
            workbook.write(outputStream);
        }
        workbook.close();
    }

    // Внутренний класс для хранения результатов студентов
    static class StudentResultData {
        private String code;
        private String className;
        private String subject;
        private String date;
        private String masteryLevel;
        private String section1Percent;
        private String section2Percent;
        private String section3Percent;

        // Геттеры и сеттеры
        public String getCode() { return code; }
        public void setCode(String code) { this.code = code; }

        public String getClassName() { return className; }
        public void setClassName(String className) { this.className = className; }

        public String getSubject() { return subject; }
        public void setSubject(String subject) { this.subject = subject; }

        public String getDate() { return date; }
        public void setDate(String date) { this.date = date; }

        public String getMasteryLevel() { return masteryLevel; }
        public void setMasteryLevel(String masteryLevel) { this.masteryLevel = masteryLevel; }

        public String getSection1Percent() { return section1Percent; }
        public void setSection1Percent(String section1Percent) { this.section1Percent = section1Percent; }

        public String getSection2Percent() { return section2Percent; }
        public void setSection2Percent(String section2Percent) { this.section2Percent = section2Percent; }

        public String getSection3Percent() { return section3Percent; }
        public void setSection3Percent(String section3Percent) { this.section3Percent = section3Percent; }

        @Override
        public String toString() {
            return String.format("%s | %s | %s | %s | %s | %s | %s | %s",
                    code, className, subject, date,
                    section1Percent, section2Percent, section3Percent, masteryLevel);
        }
    }
}