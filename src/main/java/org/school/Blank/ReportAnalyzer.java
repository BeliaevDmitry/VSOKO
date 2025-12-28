package org.school.Blank;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.stream.Collectors;

public class ReportAnalyzer {

    // ===== НАСТРОЙКИ ПУТЕЙ =====
    private static final String INPUT_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\На разбор";
    private static final String REPORTS_BASE_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы";
    private static final String FINAL_REPORT_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\История";

    private static final String FILE_PATTERN = "Сбор_данных_*";

    // ===== СТРУКТУРЫ ДАННЫХ =====
    static class AnalysisResult {
        String subject;
        String className;
        String teacher;
        String date;
        String parallel;
        int totalStudents;
        int presentStudents;
        double averageScore;
        double averagePercentage;
        Map<Integer, Double> taskSuccessRates;
        Map<Integer, Integer> taskScores;
        int maxPossibleScore;

        public AnalysisResult() {
            taskSuccessRates = new TreeMap<>();
            taskScores = new TreeMap<>();
        }
    }

    static class TeacherSummary {
        String teacherName;
        int totalClasses;
        int totalStudents;
        int totalPresentStudents;
        double averageScoreOverall;
        double averagePercentageOverall;
        Map<String, Double> classAverages = new HashMap<>();
        Map<String, AnalysisResult> classResults = new HashMap<>();

        public void addClassResult(AnalysisResult result) {
            classResults.put(result.className, result);
            classAverages.put(result.className, result.averageScore);

            totalClasses++;
            totalStudents += result.totalStudents;
            totalPresentStudents += result.presentStudents;

            // Пересчитываем среднее
            averageScoreOverall = classResults.values().stream()
                    .mapToDouble(r -> r.averageScore)
                    .average()
                    .orElse(0);

            averagePercentageOverall = classResults.values().stream()
                    .mapToDouble(r -> r.averagePercentage)
                    .average()
                    .orElse(0);
        }
    }

    static class ParallelSummary {
        String parallel;
        String subject;
        int totalClasses;
        int totalStudents;
        int totalPresentStudents;
        double averageScoreParallel;
        double averagePercentageParallel;
        Map<String, TeacherSummary> teacherSummaries = new HashMap<>();
        Map<String, AnalysisResult> classResults = new HashMap<>();

        public ParallelSummary(String parallel, String subject) {
            this.parallel = parallel;
            this.subject = subject;
        }

        public void addClassResult(AnalysisResult result) {
            classResults.put(result.className, result);

            totalClasses++;
            totalStudents += result.totalStudents;
            totalPresentStudents += result.presentStudents;

            // Обновляем статистику по учителю
            TeacherSummary teacherSummary = teacherSummaries.get(result.teacher);
            if (teacherSummary == null) {
                teacherSummary = new TeacherSummary();
                teacherSummary.teacherName = result.teacher;
                teacherSummaries.put(result.teacher, teacherSummary);
            }
            teacherSummary.addClassResult(result);

            // Пересчитываем среднее по параллели
            averageScoreParallel = classResults.values().stream()
                    .mapToDouble(r -> r.averageScore)
                    .average()
                    .orElse(0);

            averagePercentageParallel = classResults.values().stream()
                    .mapToDouble(r -> r.averagePercentage)
                    .average()
                    .orElse(0);
        }

        public Map<String, Double> getClassRankings() {
            // Исправленная версия для Java 11
            List<Map.Entry<String, AnalysisResult>> sortedEntries = new ArrayList<>(classResults.entrySet());
            sortedEntries.sort((e1, e2) -> Double.compare(e2.getValue().averageScore, e1.getValue().averageScore));

            Map<String, Double> rankings = new LinkedHashMap<>();
            for (Map.Entry<String, AnalysisResult> entry : sortedEntries) {
                rankings.put(entry.getKey(), entry.getValue().averageScore);
            }
            return rankings;
        }
    }

    public static void main(String[] args) {
        try {
            System.out.println("=== СИСТЕМА АНАЛИЗА ОТЧЁТОВ ===");
            System.out.println("Папка для разбора: " + INPUT_FOLDER);

            // 1. Поиск файлов отчётов
            List<Path> reportFiles = findReportFiles(INPUT_FOLDER);

            if (reportFiles.isEmpty()) {
                System.out.println("Файлы отчётов не найдены в папке: " + INPUT_FOLDER);
                return;
            }

            System.out.println("\nНайдено файлов: " + reportFiles.size());

            // 2. Анализ каждого файла
            Map<String, ParallelSummary> parallelSummaries = new TreeMap<>();

            for (Path file : reportFiles) {
                System.out.println("\nАнализ файла: " + file.getFileName());

                try {
                    AnalysisResult result = analyzeReportFile(file);
                    if (result != null) {
                        String parallelKey = result.subject + "_" + result.parallel;

                        ParallelSummary parallelSummary = parallelSummaries.get(parallelKey);
                        if (parallelSummary == null) {
                            parallelSummary = new ParallelSummary(result.parallel, result.subject);
                            parallelSummaries.put(parallelKey, parallelSummary);
                        }
                        parallelSummary.addClassResult(result);

                        // Перенос файла в соответствующую папку предмета
                        moveFileToSubjectFolder(file, result.subject);
                    }
                } catch (Exception e) {
                    System.err.println("Ошибка при анализе файла " + file + ": " + e.getMessage());
                    e.printStackTrace();
                }
            }

            // 3. Создание итоговых отчётов
            System.out.println("\n=== СОЗДАНИЕ ИТОГОВЫХ ОТЧЁТОВ ===");

            for (ParallelSummary parallelSummary : parallelSummaries.values()) {
                createParallelReport(parallelSummary);
                createTeacherReports(parallelSummary);
                createCharts(parallelSummary);
            }

            // 4. Создание сводного отчёта по всем параллелям
            if (!parallelSummaries.isEmpty()) {
                createConsolidatedReport(parallelSummaries);
            }

            System.out.println("\n=== АНАЛИЗ ЗАВЕРШЁН ===");
            System.out.println("Обработано параллелей: " + parallelSummaries.size());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<Path> findReportFiles(String folderPath) throws IOException {
        List<Path> files = new ArrayList<>();
        Path folder = Paths.get(folderPath);

        if (Files.exists(folder) && Files.isDirectory(folder)) {
            try (DirectoryStream<Path> stream = Files.newDirectoryStream(folder,
                    path -> path.getFileName().toString().startsWith("Сбор_данных_")
                            && path.toString().endsWith(".xlsx"))) {
                for (Path path : stream) {
                    files.add(path);
                }
            }
        }

        return files;
    }

    private static AnalysisResult analyzeReportFile(Path filePath) throws IOException {
        try (FileInputStream file = new FileInputStream(filePath.toFile());
             Workbook workbook = new XSSFWorkbook(file)) {

            AnalysisResult result = new AnalysisResult();

            // 1. Чтение информации из листа "Информация"
            Sheet infoSheet = workbook.getSheet("Информация");
            if (infoSheet == null) {
                throw new IllegalArgumentException("Лист 'Информация' не найден");
            }

            // Извлекаем базовую информацию
            result.teacher = getCellValue(infoSheet, 0, 1);
            result.date = getCellValue(infoSheet, 1, 1);
            result.subject = getCellValue(infoSheet, 2, 1);
            result.className = getCellValue(infoSheet, 3, 1);
            result.parallel = extractParallel(result.className);

            System.out.println("  Класс: " + result.className);
            System.out.println("  Учитель: " + result.teacher);
            System.out.println("  Предмет: " + result.subject);

            // 2. Анализ данных из листа "Сбор информации"
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            if (dataSheet == null) {
                throw new IllegalArgumentException("Лист 'Сбор информации' не найден");
            }

            // Находим строки с максимальными баллами (третья строка)
            Row maxScoresRow = dataSheet.getRow(2);
            if (maxScoresRow == null) {
                throw new IllegalArgumentException("Строка с максимальными баллами не найдена");
            }

            // Собираем максимальные баллы за задания
            List<Integer> maxScores = new ArrayList<>();
            int taskCol = 4;

            while (true) {
                Cell maxScoreCell = maxScoresRow.getCell(taskCol);
                if (maxScoreCell == null || maxScoreCell.getCellType() == CellType.BLANK) {
                    break;
                }

                try {
                    int maxScore = (int) maxScoreCell.getNumericCellValue();
                    maxScores.add(maxScore);
                    taskCol++;
                } catch (Exception e) {
                    break;
                }
            }

            result.maxPossibleScore = maxScores.stream().mapToInt(Integer::intValue).sum();
            System.out.println("  Макс. возможный балл: " + result.maxPossibleScore);

            // 3. Анализ данных учеников
            int firstStudentRow = 3;
            int studentCount = 0;
            int presentCount = 0;
            double totalScore = 0;
            Map<Integer, Double> taskTotals = new HashMap<>();
            Map<Integer, Integer> taskMaxScores = new HashMap<>();

            // Инициализируем структуры для заданий
            for (int i = 0; i < maxScores.size(); i++) {
                taskTotals.put(i + 1, 0.0);
                taskMaxScores.put(i + 1, maxScores.get(i));
            }

            for (int rowNum = firstStudentRow; rowNum <= dataSheet.getLastRowNum(); rowNum++) {
                Row row = dataSheet.getRow(rowNum);
                if (row == null) continue;

                Cell fioCell = row.getCell(1);
                if (fioCell == null || fioCell.getCellType() == CellType.BLANK) {
                    continue;
                }

                String fio = getCellValue(fioCell);
                if (fio == null || fio.trim().isEmpty()) {
                    continue;
                }

                studentCount++;

                // Проверяем присутствие
                Cell presenceCell = row.getCell(2);
                String presence = getCellValue(presenceCell);

                if (presence != null && "Не был".equalsIgnoreCase(presence.trim())) {
                    continue;
                }

                presentCount++;

                // Считаем баллы по заданиям
                double studentScore = 0;

                for (int taskNum = 1; taskNum <= maxScores.size(); taskNum++) {
                    Cell scoreCell = row.getCell(3 + taskNum);
                    if (scoreCell != null && scoreCell.getCellType() == CellType.NUMERIC) {
                        double taskScore = scoreCell.getNumericCellValue();
                        studentScore += taskScore;

                        Double currentTotal = taskTotals.get(taskNum);
                        taskTotals.put(taskNum, currentTotal + taskScore);
                    }
                }

                totalScore += studentScore;
            }

            result.totalStudents = studentCount;
            result.presentStudents = presentCount;

            if (presentCount > 0) {
                result.averageScore = totalScore / presentCount;
                result.averagePercentage = (result.averageScore / result.maxPossibleScore) * 100;

                // Рассчитываем успешность по заданиям
                for (int taskNum = 1; taskNum <= maxScores.size(); taskNum++) {
                    double taskAvg = taskTotals.get(taskNum) / presentCount;
                    double taskSuccessRate = (taskAvg / taskMaxScores.get(taskNum)) * 100;

                    result.taskScores.put(taskNum, (int) Math.round(taskAvg));
                    result.taskSuccessRates.put(taskNum, taskSuccessRate);
                }
            }

            System.out.println("  Всего учеников: " + studentCount);
            System.out.println("  Присутствовало: " + presentCount);
            System.out.printf("  Средний балл: %.2f\n", result.averageScore);
            System.out.printf("  Средний процент: %.1f%%\n", result.averagePercentage);

            return result;

        } catch (Exception e) {
            System.err.println("Ошибка при анализе файла " + filePath + ": " + e.getMessage());
            throw e;
        }
    }

    private static String extractParallel(String className) {
        if (className == null || className.isEmpty()) {
            return "Неизвестно";
        }

        // Ищем цифры в начале строки
        StringBuilder parallel = new StringBuilder();
        for (char c : className.toCharArray()) {
            if (Character.isDigit(c)) {
                parallel.append(c);
            } else if (parallel.length() > 0) {
                break;
            }
        }

        return parallel.length() > 0 ? parallel.toString() : "Неизвестно";
    }

    private static void moveFileToSubjectFolder(Path file, String subject) {
        try {
            Path subjectFolder = Paths.get(REPORTS_BASE_FOLDER, subject, "Отчёты");
            Files.createDirectories(subjectFolder);

            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String originalName = file.getFileName().toString();
            String newFileName = "Обработанный_" + originalName.replace(".xlsx",
                    "_" + timestamp + ".xlsx");

            Path destination = subjectFolder.resolve(newFileName);

            Files.copy(file, destination, StandardCopyOption.REPLACE_EXISTING);
            Files.delete(file);

            System.out.println("  Файл перемещён в: " + destination);

        } catch (Exception e) {
            System.err.println("Ошибка при перемещении файла: " + e.getMessage());
        }
    }

    private static void createParallelReport(ParallelSummary summary) {
        try {
            Path reportFolder = Paths.get(FINAL_REPORT_FOLDER, "Анализ", summary.parallel + " класс");
            Files.createDirectories(reportFolder);

            String fileName = String.format("Отчёт_%s_%s_%s.xlsx",
                    summary.subject, summary.parallel,
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")));

            Path filePath = reportFolder.resolve(fileName);

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Map<String, XSSFCellStyle> styles = createStyles(workbook);

                // ===== ЛИСТ 1: Сводка по параллели =====
                XSSFSheet summarySheet = workbook.createSheet("Сводка по параллели");

                // Заголовок
                Row titleRow = summarySheet.createRow(0);
                Cell titleCell = titleRow.createCell(0);
                titleCell.setCellValue(String.format("СВОДНЫЙ АНАЛИЗ ПО ПАРАЛЛЕЛИ %s КЛАСС (%s)",
                        summary.parallel, summary.subject));
                titleCell.setCellStyle(styles.get("title"));
                summarySheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));

                Row subTitleRow = summarySheet.createRow(1);
                subTitleRow.createCell(0).setCellValue("Дата формирования: " +
                        LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm")));
                subTitleRow.getCell(0).setCellStyle(styles.get("subtitle"));
                summarySheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 8));

                // Статистика по параллели
                int rowNum = 3;

                Row statsRow1 = summarySheet.createRow(rowNum++);
                statsRow1.createCell(0).setCellValue("Всего классов:");
                statsRow1.createCell(1).setCellValue(summary.totalClasses);
                statsRow1.getCell(0).setCellStyle(styles.get("header"));
                statsRow1.getCell(1).setCellStyle(styles.get("data"));

                Row statsRow2 = summarySheet.createRow(rowNum++);
                statsRow2.createCell(0).setCellValue("Всего учеников:");
                statsRow2.createCell(1).setCellValue(summary.totalStudents);

                Row statsRow3 = summarySheet.createRow(rowNum++);
                statsRow3.createCell(0).setCellValue("Присутствовало:");
                statsRow3.createCell(1).setCellValue(summary.totalPresentStudents);

                Row statsRow4 = summarySheet.createRow(rowNum++);
                statsRow4.createCell(0).setCellValue("Средний балл по параллели:");
                statsRow4.createCell(1).setCellValue(summary.averageScoreParallel);
                statsRow4.getCell(1).setCellStyle(styles.get("data_bold"));

                Row statsRow5 = summarySheet.createRow(rowNum++);
                statsRow5.createCell(0).setCellValue("Средний % выполнения:");
                Cell percentCell = statsRow5.createCell(1);
                percentCell.setCellValue(summary.averagePercentageParallel / 100);
                percentCell.setCellStyle(styles.get("percent_cell"));

                rowNum++;

                // ===== ТАБЛИЦА: Рейтинг классов =====
                Row tableHeaderRow = summarySheet.createRow(rowNum++);
                String[] headers = {"Место", "Класс", "Учитель", "Учеников", "Присут.",
                        "Ср. балл", "% вып.", "Отклонение от среднего"};

                for (int i = 0; i < headers.length; i++) {
                    Cell cell = tableHeaderRow.createCell(i);
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(styles.get("table_header"));
                }

                // Заполняем данные по классам
                int place = 1;
                Map<String, Double> rankings = summary.getClassRankings();

                for (Map.Entry<String, Double> entry : rankings.entrySet()) {
                    String className = entry.getKey();
                    AnalysisResult result = summary.classResults.get(className);

                    Row dataRow = summarySheet.createRow(rowNum++);

                    dataRow.createCell(0).setCellValue(place++);
                    dataRow.createCell(1).setCellValue(className);
                    dataRow.createCell(2).setCellValue(result.teacher);
                    dataRow.createCell(3).setCellValue(result.totalStudents);
                    dataRow.createCell(4).setCellValue(result.presentStudents);
                    dataRow.createCell(5).setCellValue(result.averageScore);

                    Cell classPercentCell = dataRow.createCell(6);
                    classPercentCell.setCellValue(result.averagePercentage / 100);
                    classPercentCell.setCellStyle(styles.get("percent_cell"));

                    // Отклонение от среднего по параллели
                    double deviation = result.averageScore - summary.averageScoreParallel;
                    Cell deviationCell = dataRow.createCell(7);
                    deviationCell.setCellValue(deviation);

                    if (deviation > 0) {
                        deviationCell.setCellStyle(styles.get("positive"));
                    } else if (deviation < 0) {
                        deviationCell.setCellStyle(styles.get("negative"));
                    } else {
                        deviationCell.setCellStyle(styles.get("data"));
                    }
                }

                // Настраиваем ширину колонок
                for (int i = 0; i < headers.length; i++) {
                    summarySheet.autoSizeColumn(i);
                }

                // Сохраняем файл
                try (FileOutputStream fileOut = new FileOutputStream(filePath.toFile())) {
                    workbook.write(fileOut);
                }

                System.out.println("\nСоздан отчёт по параллели: " + filePath);

            }

        } catch (Exception e) {
            System.err.println("Ошибка при создании отчёта по параллели: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static Map<String, XSSFCellStyle> createStyles(XSSFWorkbook workbook) {
        Map<String, XSSFCellStyle> styles = new HashMap<>();

        // Стиль заголовка
        XSSFCellStyle titleStyle = workbook.createCellStyle();
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 16);
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        styles.put("title", titleStyle);

        // Стиль подзаголовка
        XSSFCellStyle subtitleStyle = workbook.createCellStyle();
        Font subtitleFont = workbook.createFont();
        subtitleFont.setItalic(true);
        subtitleFont.setFontHeightInPoints((short) 12);
        subtitleStyle.setFont(subtitleFont);
        subtitleStyle.setAlignment(HorizontalAlignment.CENTER);
        styles.put("subtitle", subtitleStyle);

        // Стиль заголовка таблицы
        XSSFCellStyle tableHeaderStyle = workbook.createCellStyle();
        Font tableHeaderFont = workbook.createFont();
        tableHeaderFont.setBold(true);
        tableHeaderStyle.setFont(tableHeaderFont);
        tableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        tableHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        tableHeaderStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        tableHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        tableHeaderStyle.setBorderTop(BorderStyle.THIN);
        tableHeaderStyle.setBorderBottom(BorderStyle.THIN);
        tableHeaderStyle.setBorderLeft(BorderStyle.THIN);
        tableHeaderStyle.setBorderRight(BorderStyle.THIN);
        styles.put("table_header", tableHeaderStyle);

        // Стиль заголовка
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        styles.put("header", headerStyle);

        // Стиль данных
        XSSFCellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        styles.put("data", dataStyle);

        // Стиль данных жирный
        XSSFCellStyle dataBoldStyle = workbook.createCellStyle();
        dataBoldStyle.cloneStyleFrom(dataStyle);
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        dataBoldStyle.setFont(boldFont);
        styles.put("data_bold", dataBoldStyle);

        // Стиль ячейки с процентом
        XSSFCellStyle percentCellStyle = workbook.createCellStyle();
        percentCellStyle.cloneStyleFrom(dataStyle);
        percentCellStyle.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        styles.put("percent_cell", percentCellStyle);

        // Стиль положительных значений
        XSSFCellStyle positiveStyle = workbook.createCellStyle();
        positiveStyle.cloneStyleFrom(dataStyle);
        positiveStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        positiveStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font positiveFont = workbook.createFont();
        positiveFont.setColor(IndexedColors.DARK_GREEN.getIndex());
        positiveStyle.setFont(positiveFont);
        styles.put("positive", positiveStyle);

        // Стиль отрицательных значений
        XSSFCellStyle negativeStyle = workbook.createCellStyle();
        negativeStyle.cloneStyleFrom(dataStyle);
        negativeStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
        negativeStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font negativeFont = workbook.createFont();
        negativeFont.setColor(IndexedColors.DARK_RED.getIndex());
        negativeStyle.setFont(negativeFont);
        styles.put("negative", negativeStyle);

        return styles;
    }

    private static String getCellValue(Sheet sheet, int row, int col) {
        Row sheetRow = sheet.getRow(row);
        if (sheetRow == null) return "";

        Cell cell = sheetRow.getCell(col);
        return getCellValue(cell);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        // Используем старый синтаксис switch для совместимости с Java 11
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return String.valueOf((int) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e2) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }

    // Добавьте остальные методы сюда...

    static class TeacherRating {
        String teacherName;
        String parallel;
        int classCount;
        double avgScore;
        double avgPercent;
        double efficiencyScore;

        public TeacherRating(String teacherName, String parallel, int classCount,
                             double avgScore, double avgPercent, double efficiencyScore) {
            this.teacherName = teacherName;
            this.parallel = parallel;
            this.classCount = classCount;
            this.avgScore = avgScore;
            this.avgPercent = avgPercent;
            this.efficiencyScore = efficiencyScore;
        }
    }

    private static String generateRecommendation(AnalysisResult result, ParallelSummary parallel) {
        double deviation = result.averageScore - parallel.averageScoreParallel;
        double percent = result.averagePercentage;

        if (percent >= 85) {
            return "Отличные результаты! Можно давать усложнённые задания.";
        } else if (percent >= 70) {
            return "Хорошие результаты. Рекомендуется работа над ошибками.";
        } else if (percent >= 50) {
            return "Удовлетворительные результаты. Требуется дополнительная работа.";
        } else if (deviation > 0) {
            return "Результат выше среднего по параллели, но низкий процент.";
        } else if (deviation < -1) {
            return "Требуется срочное вмешательство! Результат значительно ниже среднего.";
        } else {
            return "Требуется индивидуальная работа с отстающими учениками.";
        }
    }

    private static void createTeacherReports(ParallelSummary summary) {
        for (TeacherSummary teacher : summary.teacherSummaries.values()) {
            try {
                // Создаём папку для отчётов учителя
                Path teacherFolder = Paths.get(FINAL_REPORT_FOLDER, "Анализ",
                        summary.parallel + " класс", "Учителя");
                Files.createDirectories(teacherFolder);

                // Формируем имя файла
                String safeTeacherName = teacher.teacherName.replaceAll("[\\\\/:*?\"<>|]", "_");
                String fileName = String.format("Отчёт_учителя_%s_%s_%s.xlsx",
                        safeTeacherName, summary.subject,
                        LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd")));

                Path filePath = teacherFolder.resolve(fileName);

                try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                    Map<String, XSSFCellStyle> styles = createStyles(workbook);

                    XSSFSheet sheet = workbook.createSheet("Отчёт учителя");

                    // Заголовок
                    Row titleRow = sheet.createRow(0);
                    titleRow.createCell(0).setCellValue(
                            String.format("ОТЧЁТ УЧИТЕЛЯ: %s", teacher.teacherName));
                    titleRow.getCell(0).setCellStyle(styles.get("title"));
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));

                    Row subTitleRow = sheet.createRow(1);
                    subTitleRow.createCell(0).setCellValue(
                            String.format("Предмет: %s, Параллель: %s класс",
                                    summary.subject, summary.parallel));
                    subTitleRow.getCell(0).setCellStyle(styles.get("subtitle"));
                    sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 6));

                    Row dateRow = sheet.createRow(2);
                    dateRow.createCell(0).setCellValue("Дата формирования: " +
                            LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm")));

                    // Статистика учителя
                    int rowNum = 4;

                    Row statRow1 = sheet.createRow(rowNum++);
                    statRow1.createCell(0).setCellValue("Количество классов:");
                    statRow1.createCell(1).setCellValue(teacher.totalClasses);
                    statRow1.getCell(0).setCellStyle(styles.get("header"));
                    statRow1.getCell(1).setCellStyle(styles.get("data"));

                    Row statRow2 = sheet.createRow(rowNum++);
                    statRow2.createCell(0).setCellValue("Общее количество учеников:");
                    statRow2.createCell(1).setCellValue(teacher.totalStudents);

                    Row statRow3 = sheet.createRow(rowNum++);
                    statRow3.createCell(0).setCellValue("Присутствовало:");
                    statRow3.createCell(1).setCellValue(teacher.totalPresentStudents);

                    Row statRow4 = sheet.createRow(rowNum++);
                    statRow4.createCell(0).setCellValue("Средний балл (по классам):");
                    statRow4.createCell(1).setCellValue(teacher.averageScoreOverall);
                    statRow4.getCell(1).setCellStyle(styles.get("data_bold"));

                    Row statRow5 = sheet.createRow(rowNum++);
                    statRow5.createCell(0).setCellValue("Средний % выполнения:");
                    Cell percentCell = statRow5.createCell(1);
                    percentCell.setCellValue(teacher.averagePercentageOverall / 100);
                    percentCell.setCellStyle(styles.get("percent_cell"));

                    // Сравнение с параллелью
                    Row statRow6 = sheet.createRow(rowNum++);
                    statRow6.createCell(0).setCellValue("Сравнение с параллелью:");
                    double diff = teacher.averageScoreOverall - summary.averageScoreParallel;
                    Cell diffCell = statRow6.createCell(1);

                    if (diff > 0) {
                        diffCell.setCellValue(String.format("+%.2f (выше среднего)", diff));
                        diffCell.setCellStyle(styles.get("positive"));
                    } else if (diff < 0) {
                        diffCell.setCellValue(String.format("%.2f (ниже среднего)", diff));
                        diffCell.setCellStyle(styles.get("negative"));
                    } else {
                        diffCell.setCellValue("На уровне среднего");
                        diffCell.setCellStyle(styles.get("data"));
                    }

                    rowNum++; // Пустая строка

                    // Таблица классов учителя
                    Row tableHeader = sheet.createRow(rowNum++);
                    String[] headers = {"Класс", "Учеников", "Присут.", "Ср. балл", "% вып.",
                            "Отклонение от среднего", "Рекомендации"};

                    for (int i = 0; i < headers.length; i++) {
                        Cell cell = tableHeader.createCell(i);
                        cell.setCellValue(headers[i]);
                        cell.setCellStyle(styles.get("table_header"));
                    }

                    // Сортируем классы по среднему баллу
                    List<Map.Entry<String, AnalysisResult>> sortedClasses = new ArrayList<>(teacher.classResults.entrySet());
                    sortedClasses.sort((e1, e2) -> Double.compare(
                            e2.getValue().averageScore,
                            e1.getValue().averageScore));

                    for (Map.Entry<String, AnalysisResult> entry : sortedClasses) {
                        AnalysisResult result = entry.getValue();

                        Row classRow = sheet.createRow(rowNum++);

                        classRow.createCell(0).setCellValue(result.className);
                        classRow.createCell(1).setCellValue(result.totalStudents);
                        classRow.createCell(2).setCellValue(result.presentStudents);
                        classRow.createCell(3).setCellValue(result.averageScore);

                        Cell classPercentCell = classRow.createCell(4);
                        classPercentCell.setCellValue(result.averagePercentage / 100);
                        classPercentCell.setCellStyle(styles.get("percent_cell"));

                        // Отклонение от среднего по параллели
                        double deviation = result.averageScore - summary.averageScoreParallel;
                        Cell deviationCell = classRow.createCell(5);
                        deviationCell.setCellValue(deviation);

                        if (deviation > 0) {
                            deviationCell.setCellStyle(styles.get("positive"));
                        } else if (deviation < 0) {
                            deviationCell.setCellStyle(styles.get("negative"));
                        } else {
                            deviationCell.setCellStyle(styles.get("data"));
                        }

                        // Рекомендации
                        Cell recommendationCell = classRow.createCell(6);
                        String recommendation = generateRecommendation(result, summary);
                        recommendationCell.setCellValue(recommendation);

                        // Автоматически настраиваем ширину для рекомендаций
                        sheet.autoSizeColumn(6);
                    }

                    // Анализ заданий для учителя
                    rowNum += 2;
                    Row taskHeader = sheet.createRow(rowNum++);
                    taskHeader.createCell(0).setCellValue("Анализ выполнения заданий:");
                    taskHeader.getCell(0).setCellStyle(styles.get("header"));
                    sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 5));

                    // Заголовки таблицы заданий
                    Row taskTableHeader = sheet.createRow(rowNum++);
                    String[] taskHeaders = {"№ задания", "Ср. % выполнения",
                            "Лучший класс", "Худший класс", "Проблемные задания"};

                    for (int i = 0; i < taskHeaders.length; i++) {
                        Cell cell = taskTableHeader.createCell(i);
                        cell.setCellValue(taskHeaders[i]);
                        cell.setCellStyle(styles.get("table_header"));
                    }

                    // Анализируем задания
                    if (!teacher.classResults.isEmpty()) {
                        AnalysisResult sampleResult = teacher.classResults.values().iterator().next();
                        int totalTasks = sampleResult.taskSuccessRates.size();

                        for (int taskNum = 1; taskNum <= totalTasks; taskNum++) {
                            Row taskRow = sheet.createRow(rowNum++);

                            // Собираем статистику по заданию для классов учителя
                            double totalSuccessRate = 0;
                            String bestClass = "";
                            double bestRate = 0;
                            String worstClass = "";
                            double worstRate = 100;
                            int classCount = 0;

                            for (Map.Entry<String, AnalysisResult> entry : teacher.classResults.entrySet()) {
                                AnalysisResult result = entry.getValue();
                                Double successRate = result.taskSuccessRates.get(taskNum);

                                if (successRate != null) {
                                    totalSuccessRate += successRate;
                                    classCount++;

                                    if (successRate > bestRate) {
                                        bestRate = successRate;
                                        bestClass = entry.getKey();
                                    }
                                    if (successRate < worstRate) {
                                        worstRate = successRate;
                                        worstClass = entry.getKey();
                                    }
                                }
                            }

                            double avgSuccessRate = classCount > 0 ? totalSuccessRate / classCount : 0;

                            taskRow.createCell(0).setCellValue(taskNum);

                            Cell taskPercentCell = taskRow.createCell(1);
                            taskPercentCell.setCellValue(avgSuccessRate / 100);
                            taskPercentCell.setCellStyle(styles.get("percent_cell"));

                            taskRow.createCell(2).setCellValue(bestClass);
                            taskRow.createCell(3).setCellValue(worstClass);

                            // Выделяем проблемные задания
                            Cell problemCell = taskRow.createCell(4);
                            if (avgSuccessRate < 50) {
                                problemCell.setCellValue("⚠ Требует внимания");
                                problemCell.setCellStyle(styles.get("negative"));
                            } else if (avgSuccessRate < 70) {
                                problemCell.setCellValue("Норма");
                                problemCell.setCellStyle(styles.get("data"));
                            } else {
                                problemCell.setCellValue("Хорошо");
                                problemCell.setCellStyle(styles.get("positive"));
                            }
                        }
                    }

                    // Настраиваем ширину всех колонок
                    for (int i = 0; i < headers.length; i++) {
                        sheet.autoSizeColumn(i);
                    }

                    // Сохраняем файл
                    try (FileOutputStream fileOut = new FileOutputStream(filePath.toFile())) {
                        workbook.write(fileOut);
                    }

                    System.out.println("Создан отчёт для учителя: " + teacher.teacherName);

                }

            } catch (Exception e) {
                System.err.println("Ошибка при создании отчёта учителя " +
                        teacher.teacherName + ": " + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    private static void createCharts(ParallelSummary summary) {
        try {
            Path chartsFolder = Paths.get(FINAL_REPORT_FOLDER, "Анализ",
                    summary.parallel + " класс", "Диаграммы");
            Files.createDirectories(chartsFolder);

            // 1. Диаграмма рейтинга классов
            createBarChart(summary, chartsFolder);

            // 2. Круговая диаграмма распределения
            createPieChart(summary, chartsFolder);

            // 3. График успешности по заданиям
            createTaskChart(summary, chartsFolder);

        } catch (Exception e) {
            System.err.println("Ошибка при создании диаграмм: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void createBarChart(ParallelSummary summary, Path folder) {
        try {
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();

            // Сортируем классы по среднему баллу
            Map<String, Double> rankings = summary.getClassRankings();

            for (Map.Entry<String, Double> entry : rankings.entrySet()) {
                dataset.addValue(entry.getValue(), "Средний балл", entry.getKey());
            }

            JFreeChart chart = ChartFactory.createBarChart(
                    String.format("Рейтинг классов (%s класс, %s)",
                            summary.parallel, summary.subject),
                    "Класс",
                    "Средний балл",
                    dataset,
                    PlotOrientation.VERTICAL,
                    true,
                    true,
                    false
            );

            Path chartPath = folder.resolve("1_рейтинг_классов.png");
            ChartUtils.saveChartAsPNG(chartPath.toFile(), chart, 1000, 600);

            System.out.println("Создана диаграмма рейтинга классов");

        } catch (Exception e) {
            System.err.println("Ошибка при создании bar chart: " + e.getMessage());
        }
    }

    private static void createPieChart(ParallelSummary summary, Path folder) {
        try {
            DefaultPieDataset dataset = new DefaultPieDataset();

            // Группируем по диапазонам процентов
            int excellent = 0; // 85-100%
            int good = 0;      // 70-84%
            int satisfactory = 0; // 50-69%
            int unsatisfactory = 0; // 0-49%

            for (AnalysisResult result : summary.classResults.values()) {
                double percent = result.averagePercentage;
                if (percent >= 85) {
                    excellent++;
                } else if (percent >= 70) {
                    good++;
                } else if (percent >= 50) {
                    satisfactory++;
                } else {
                    unsatisfactory++;
                }
            }

            if (excellent > 0) dataset.setValue("Отлично (85-100%)", excellent);
            if (good > 0) dataset.setValue("Хорошо (70-84%)", good);
            if (satisfactory > 0) dataset.setValue("Удовл. (50-69%)", satisfactory);
            if (unsatisfactory > 0) dataset.setValue("Неудовл. (<50%)", unsatisfactory);

            JFreeChart chart = ChartFactory.createPieChart(
                    String.format("Распределение классов по успеваемости (%s класс)",
                            summary.parallel),
                    dataset,
                    true,
                    true,
                    false
            );

            Path chartPath = folder.resolve("2_распределение_успеваемости.png");
            ChartUtils.saveChartAsPNG(chartPath.toFile(), chart, 800, 600);

            System.out.println("Создана круговая диаграмма распределения");

        } catch (Exception e) {
            System.err.println("Ошибка при создании pie chart: " + e.getMessage());
        }
    }

    private static void createTaskChart(ParallelSummary summary, Path folder) {
        try {
            DefaultCategoryDataset dataset = new DefaultCategoryDataset();

            // Берем первый класс для структуры заданий
            AnalysisResult sampleResult = null;
            for (AnalysisResult result : summary.classResults.values()) {
                sampleResult = result;
                break;
            }

            if (sampleResult != null) {
                // Рассчитываем средний процент по каждому заданию
                for (int taskNum = 1; taskNum <= sampleResult.taskSuccessRates.size(); taskNum++) {
                    double totalRate = 0;
                    int count = 0;

                    for (AnalysisResult result : summary.classResults.values()) {
                        Double rate = result.taskSuccessRates.get(taskNum);
                        if (rate != null) {
                            totalRate += rate;
                            count++;
                        }
                    }

                    if (count > 0) {
                        double avgRate = totalRate / count;
                        dataset.addValue(avgRate, "% выполнения", "Задание " + taskNum);
                    }
                }

                JFreeChart chart = ChartFactory.createLineChart(
                        String.format("Успешность выполнения заданий (%s класс)",
                                summary.parallel),
                        "Номер задания",
                        "% выполнения",
                        dataset,
                        PlotOrientation.VERTICAL,
                        true,
                        true,
                        false
                );

                Path chartPath = folder.resolve("3_успешность_заданий.png");
                ChartUtils.saveChartAsPNG(chartPath.toFile(), chart, 1200, 600);

                System.out.println("Создан график успешности по заданиям");
            }

        } catch (Exception e) {
            System.err.println("Ошибка при создании task chart: " + e.getMessage());
        }
    }

    private static void createConsolidatedReport(Map<String, ParallelSummary> parallelSummaries) {
        try {
            // Создаём папку для сводных отчётов
            Path consolidatedFolder = Paths.get(FINAL_REPORT_FOLDER, "Сводные отчёты");
            Files.createDirectories(consolidatedFolder);

            String fileName = String.format("Сводный_отчёт_ВСЕ_ПАРАЛЛЕЛИ_%s.xlsx",
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmm")));

            Path filePath = consolidatedFolder.resolve(fileName);

            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Map<String, XSSFCellStyle> styles = createStyles(workbook);

                // ===== ЛИСТ 1: Сводка по всем параллелям =====
                XSSFSheet sheet = workbook.createSheet("Сводный отчёт");

                // Заголовок
                Row titleRow = sheet.createRow(0);
                titleRow.createCell(0).setCellValue("СВОДНЫЙ АНАЛИЗ РАБОТ ПО ВСЕМ ПАРАЛЛЕЛЯМ");
                titleRow.getCell(0).setCellStyle(styles.get("title"));
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 8));

                Row subTitleRow = sheet.createRow(1);
                subTitleRow.createCell(0).setCellValue("Предмет: История");
                subTitleRow.getCell(0).setCellStyle(styles.get("subtitle"));
                sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 8));

                Row dateRow = sheet.createRow(2);
                dateRow.createCell(0).setCellValue("Дата формирования: " +
                        LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm")));

                // Таблица сводных данных
                int rowNum = 4;
                Row headerRow = sheet.createRow(rowNum++);

                String[] headers = {"Параллель", "Кол-во классов", "Всего учеников",
                        "Присутствовало", "Ср. балл", "Ср. %", "Лучший класс",
                        "Лучший балл", "Лучший учитель", "Эффективность"};

                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(styles.get("table_header"));
                }

                // Сортируем параллели по номеру
                List<ParallelSummary> sortedSummaries = new ArrayList<>(parallelSummaries.values());
                sortedSummaries.sort((s1, s2) -> {
                    try {
                        int p1 = Integer.parseInt(s1.parallel);
                        int p2 = Integer.parseInt(s2.parallel);
                        return Integer.compare(p1, p2);
                    } catch (NumberFormatException e) {
                        return s1.parallel.compareTo(s2.parallel);
                    }
                });

                // Заполняем данные по параллелям
                for (ParallelSummary summary : sortedSummaries) {
                    Row dataRow = sheet.createRow(rowNum++);

                    dataRow.createCell(0).setCellValue(summary.parallel + " класс");
                    dataRow.createCell(1).setCellValue(summary.totalClasses);
                    dataRow.createCell(2).setCellValue(summary.totalStudents);
                    dataRow.createCell(3).setCellValue(summary.totalPresentStudents);
                    dataRow.createCell(4).setCellValue(summary.averageScoreParallel);

                    Cell percentCell = dataRow.createCell(5);
                    percentCell.setCellValue(summary.averagePercentageParallel / 100);
                    percentCell.setCellStyle(styles.get("percent_cell"));

                    // Находим лучший класс
                    AnalysisResult bestResult = null;
                    String bestClassName = "";
                    double bestScore = -1;

                    for (Map.Entry<String, AnalysisResult> entry : summary.classResults.entrySet()) {
                        if (entry.getValue().averageScore > bestScore) {
                            bestScore = entry.getValue().averageScore;
                            bestClassName = entry.getKey();
                            bestResult = entry.getValue();
                        }
                    }

                    if (bestResult != null) {
                        dataRow.createCell(6).setCellValue(bestClassName);
                        dataRow.createCell(7).setCellValue(bestResult.averageScore);
                        dataRow.createCell(8).setCellValue(bestResult.teacher);

                        // Оценка эффективности параллели
                        Cell efficiencyCell = dataRow.createCell(9);
                        double attendanceRate = (double) summary.totalPresentStudents / summary.totalStudents * 100;

                        if (summary.averagePercentageParallel >= 70 && attendanceRate >= 90) {
                            efficiencyCell.setCellValue("Высокая");
                            efficiencyCell.setCellStyle(styles.get("positive"));
                        } else if (summary.averagePercentageParallel >= 50 && attendanceRate >= 80) {
                            efficiencyCell.setCellValue("Средняя");
                            efficiencyCell.setCellStyle(styles.get("data"));
                        } else {
                            efficiencyCell.setCellValue("Требует внимания");
                            efficiencyCell.setCellStyle(styles.get("negative"));
                        }
                    }
                }

                // Общая статистика
                rowNum += 2;
                Row totalHeader = sheet.createRow(rowNum++);
                totalHeader.createCell(0).setCellValue("ОБЩАЯ СТАТИСТИКА:");
                totalHeader.getCell(0).setCellStyle(styles.get("header"));
                sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 3));

                int totalClasses = 0;
                int totalStudents = 0;
                int totalPresent = 0;
                double totalScore = 0;
                double totalPercent = 0;

                for (ParallelSummary summary : sortedSummaries) {
                    totalClasses += summary.totalClasses;
                    totalStudents += summary.totalStudents;
                    totalPresent += summary.totalPresentStudents;
                    totalScore += summary.averageScoreParallel;
                    totalPercent += summary.averagePercentageParallel;
                }

                double overallAvgScore = sortedSummaries.isEmpty() ? 0 : totalScore / sortedSummaries.size();
                double overallAvgPercent = sortedSummaries.isEmpty() ? 0 : totalPercent / sortedSummaries.size();

                Row totalRow1 = sheet.createRow(rowNum++);
                totalRow1.createCell(0).setCellValue("Всего классов:");
                totalRow1.createCell(1).setCellValue(totalClasses);

                Row totalRow2 = sheet.createRow(rowNum++);
                totalRow2.createCell(0).setCellValue("Всего учеников:");
                totalRow2.createCell(1).setCellValue(totalStudents);

                Row totalRow3 = sheet.createRow(rowNum++);
                totalRow3.createCell(0).setCellValue("Присутствовало:");
                totalRow3.createCell(1).setCellValue(totalPresent);

                Row totalRow4 = sheet.createRow(rowNum++);
                totalRow4.createCell(0).setCellValue("Общий средний балл:");
                totalRow4.createCell(1).setCellValue(overallAvgScore);
                totalRow4.getCell(1).setCellStyle(styles.get("data_bold"));

                Row totalRow5 = sheet.createRow(rowNum++);
                totalRow5.createCell(0).setCellValue("Общий средний %:");
                Cell totalPercentCell = totalRow5.createCell(1);
                totalPercentCell.setCellValue(overallAvgPercent / 100);
                totalPercentCell.setCellStyle(styles.get("percent_cell"));

                // Настраиваем ширину колонок
                for (int i = 0; i < headers.length; i++) {
                    sheet.autoSizeColumn(i);
                }

                // Сохраняем файл
                try (FileOutputStream fileOut = new FileOutputStream(filePath.toFile())) {
                    workbook.write(fileOut);
                }

                System.out.println("\nСоздан сводный отчёт по всем параллелям: " + filePath);

            }

        } catch (Exception e) {
            System.err.println("Ошибка при создании сводного отчёта: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static double calculateEfficiencyScore(TeacherSummary teacher, ParallelSummary parallel) {
        // Рассчитываем оценку эффективности учителя (0-10)
        double score = 0;

        // 1. Средний балл относительно параллели (макс 4 балла)
        double scoreDiff = teacher.averageScoreOverall - parallel.averageScoreParallel;
        if (scoreDiff > 1) score += 4;
        else if (scoreDiff > 0.5) score += 3;
        else if (scoreDiff > 0) score += 2;
        else if (scoreDiff > -0.5) score += 1;

        // 2. Средний процент выполнения (макс 3 балла)
        if (teacher.averagePercentageOverall >= 80) score += 3;
        else if (teacher.averagePercentageOverall >= 70) score += 2;
        else if (teacher.averagePercentageOverall >= 60) score += 1;

        // 3. Посещаемость (макс 2 балла)
        double attendanceRate = (double) teacher.totalPresentStudents / teacher.totalStudents;
        if (attendanceRate >= 0.95) score += 2;
        else if (attendanceRate >= 0.90) score += 1;

        // 4. Количество классов (макс 1 балл)
        if (teacher.totalClasses >= 3) score += 1;

        return score;
    }

    private static double calculateParallelRating(ParallelSummary summary) {
        // Рассчитываем оценку параллели (0-10)
        double rating = 0;

        // 1. Средний процент выполнения (макс 4 балла)
        if (summary.averagePercentageParallel >= 80) rating += 4;
        else if (summary.averagePercentageParallel >= 70) rating += 3;
        else if (summary.averagePercentageParallel >= 60) rating += 2;
        else if (summary.averagePercentageParallel >= 50) rating += 1;

        // 2. Посещаемость (макс 3 балла)
        double attendanceRate = (double) summary.totalPresentStudents / summary.totalStudents;
        if (attendanceRate >= 0.95) rating += 3;
        else if (attendanceRate >= 0.90) rating += 2;
        else if (attendanceRate >= 0.85) rating += 1;

        // 3. Разброс результатов (макс 2 балла)
        double minScore = Double.MAX_VALUE;
        double maxScore = Double.MIN_VALUE;

        for (AnalysisResult result : summary.classResults.values()) {
            if (result.averageScore < minScore) minScore = result.averageScore;
            if (result.averageScore > maxScore) maxScore = result.averageScore;
        }

        double scoreSpread = maxScore - minScore;

        if (scoreSpread <= 1) rating += 2; // Маленький разброс - хорошо
        else if (scoreSpread <= 2) rating += 1; // Средний разброс

        // 4. Количество классов (макс 1 балл)
        if (summary.totalClasses >= 4) rating += 1;

        return rating;
    }

}