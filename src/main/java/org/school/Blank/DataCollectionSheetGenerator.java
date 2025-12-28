package org.school.Blank;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;

import java.io.*;
import java.nio.file.*;
import java.util.*;

public class DataCollectionSheetGenerator {

    // ===== НАСТРОЙКИ =====
    private static final String OUTPUT_BASE_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\История\\Формы сбора";
    private static final String EXCEL_TEMPLATE_PATH = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\Реестр контингента.xlsx";

    // Константы (можно менять под разные потребности)
    private static final String PARALLEL = "10"; // Параллель классов
    private static final String SUBJECT = "История"; // Название предмета
    private static final int MAX_TASKS = 10; // Максимальное количество заданий
    private static final int CURRENT_TASKS_COUNT = 10; // Текущее количество заданий (можно менять)
    private static final int MAX_STUDENTS = 34; // Максимальное количество учеников в классе
    private static final int VARIANTS_COUNT = 5; // Количество вариантов

    // Максимальные баллы за каждое задание (по умолчанию все по 1 баллу)
    // Пример: 1-10 задания = 1 балл, 11-12 = 3 балла, 13-19 = 1 балл
    private static final int[] MAX_SCORES_PER_TASK = {
            2, 2, 2, 2, 2,  // задания 1-5
            2, 2, 2, 2, 2,  // задания 6-10
            1, 1, 1, 1, 1,  // задания 11-15
            1, 1, 1, 1, 2,  // задания 16-20
            2, 2, 2, 2, 2,  // задания 21-25
            1, 1, 1, 1, 1,  // задания 26-30
    };

    // Цвета для подсветки
    private static final short GREEN_COLOR = IndexedColors.LIGHT_GREEN.getIndex();
    private static final short RED_COLOR = IndexedColors.RED.getIndex();
    private static final short LIGHT_YELLOW = IndexedColors.LIGHT_YELLOW.getIndex();
    private static final short LIGHT_TURQUOISE = IndexedColors.TURQUOISE.getIndex();
    private static final short LIGHT_CORNFLOWER_BLUE = IndexedColors.PALE_BLUE.getIndex();
    private static final short LIGHT_GRAY = IndexedColors.GREY_25_PERCENT.getIndex();

    public static void main(String[] args) {
        try {
            // Создаем папку для параллели, если не существует
            String parallelFolder = OUTPUT_BASE_FOLDER + "\\" + PARALLEL + " класс";
            Files.createDirectories(Paths.get(parallelFolder));

            // Читаем список классов из Excel
            List<String> classes = readClassesFromExcel();

            System.out.println("Найдены классы для обработки:");
            classes.forEach(System.out::println);

            // Создаем файлы для каждого класса
            for (String className : classes) {
                createDataCollectionFile(className, parallelFolder);
            }

            System.out.println("\nГотово! Файлы созданы в папке: " + parallelFolder);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<String> readClassesFromExcel() throws IOException {
        List<String> classes = new ArrayList<>();
        Set<String> uniqueClasses = new TreeSet<>();

        try (FileInputStream file = new FileInputStream(EXCEL_TEMPLATE_PATH);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0); // Первый лист

            // Пропускаем заголовок
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Колонка P (индекс 15) - Класс
                Cell classCell = row.getCell(15);

                if (classCell != null) {
                    String className = getCellValueAsString(classCell).trim();
                    String normalized = normalizeClassName(className);

                    // Фильтруем по параллели
                    if (normalized.startsWith(PARALLEL)) {
                        uniqueClasses.add(normalized);
                    }
                }
            }
        }

        classes.addAll(uniqueClasses);
        return classes;
    }

    private static String normalizeClassName(String className) {
        return className.trim()
                .replaceAll("\\s+", " ")
                .replaceAll("\\s*-\\s*", "-")
                .replaceAll("[^\\dА-Яа-яA-Za-z\\s-]", "");
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

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
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                return "";
        }
    }

    private static void createDataCollectionFile(String className, String parallelFolder) throws IOException {
        // Создаем имя файла
        String safeClassName = className.replaceAll("[\\\\/:*?\"<>|]", "_");
        String fileName = parallelFolder + "\\Сбор_данных_" + safeClassName + "_" + SUBJECT + ".xlsx";

        // Читаем список учеников класса
        List<Student> students = readStudentsByClass(className);

        System.out.println("\nСоздание файла для класса: " + className);
        System.out.println("Количество учеников: " + students.size());
        System.out.println("Количество заданий: " + CURRENT_TASKS_COUNT);
        System.out.println("Максимальные баллы за задания:");
        for (int i = 0; i < CURRENT_TASKS_COUNT; i++) {
            System.out.println("Задание " + (i+1) + ": " + getMaxScoreForTask(i) + " балл(ов)");
        }

        // Создаем новую рабочую книгу
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            // Создаем стили
            Map<String, XSSFCellStyle> styles = createStyles(workbook);

            // ===== ЛИСТ 1: Информация =====
            createInfoSheet(workbook, styles, className);

            // ===== ЛИСТ 2: Сбор информации =====
            createDataCollectionSheet(workbook, styles, className, students);

            // Сохраняем файл
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }

            System.out.println("Файл создан: " + fileName);
        }
    }

    private static List<Student> readStudentsByClass(String className) throws IOException {
        List<Student> students = new ArrayList<>();

        try (FileInputStream file = new FileInputStream(EXCEL_TEMPLATE_PATH);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell fioCell = row.getCell(2);
                Cell classCell = row.getCell(15);

                if (fioCell != null && classCell != null) {
                    String fio = getCellValueAsString(fioCell);
                    String studentClass = getCellValueAsString(classCell).trim();

                    if (!fio.isEmpty() && normalizeClassName(studentClass).equals(className)) {
                        students.add(new Student(fio, className));
                    }
                }
            }
        }

        // Сортируем по ФИО
        students.sort(Comparator.comparing(Student::getFio));

        return students;
    }

    private static Map<String, XSSFCellStyle> createStyles(XSSFWorkbook workbook) {
        Map<String, XSSFCellStyle> styles = new HashMap<>();

        // Стиль для заголовков
        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(LIGHT_TURQUOISE);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        styles.put("header", headerStyle);

        // Стиль для обычных ячеек
        XSSFCellStyle normalStyle = workbook.createCellStyle();
        normalStyle.setBorderTop(BorderStyle.THIN);
        normalStyle.setBorderBottom(BorderStyle.THIN);
        normalStyle.setBorderLeft(BorderStyle.THIN);
        normalStyle.setBorderRight(BorderStyle.THIN);
        normalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        styles.put("normal", normalStyle);

        // Стиль для ячеек с центрированием
        XSSFCellStyle centerStyle = workbook.createCellStyle();
        centerStyle.cloneStyleFrom(normalStyle);
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        styles.put("center", centerStyle);

        // Стиль для ячеек учителя (желтый фон)
        XSSFCellStyle teacherStyle = workbook.createCellStyle();
        teacherStyle.cloneStyleFrom(normalStyle);
        teacherStyle.setFillForegroundColor(LIGHT_YELLOW);
        teacherStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("teacher", teacherStyle);

        // Стиль для зеленого фона (Был)
        XSSFCellStyle greenStyle = workbook.createCellStyle();
        greenStyle.cloneStyleFrom(centerStyle);
        greenStyle.setFillForegroundColor(GREEN_COLOR);
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("green", greenStyle);

        // Стиль для красного фона (Не был)
        XSSFCellStyle redStyle = workbook.createCellStyle();
        redStyle.cloneStyleFrom(centerStyle);
        redStyle.setFillForegroundColor(RED_COLOR);
        redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font redFont = workbook.createFont();
        redFont.setColor(IndexedColors.WHITE.getIndex());
        redStyle.setFont(redFont);
        styles.put("red", redStyle);

        // Стиль для заголовков заданий
        XSSFCellStyle taskHeaderStyle = workbook.createCellStyle();
        taskHeaderStyle.cloneStyleFrom(centerStyle);
        taskHeaderStyle.setFillForegroundColor(LIGHT_CORNFLOWER_BLUE);
        taskHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font taskFont = workbook.createFont();
        taskFont.setBold(true);
        taskHeaderStyle.setFont(taskFont);
        styles.put("taskHeader", taskHeaderStyle);

        // Стиль для максимальных баллов
        XSSFCellStyle maxScoreStyle = workbook.createCellStyle();
        maxScoreStyle.cloneStyleFrom(centerStyle);
        maxScoreStyle.setFillForegroundColor(LIGHT_GRAY);
        maxScoreStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font scoreFont = workbook.createFont();
        scoreFont.setColor(IndexedColors.DARK_RED.getIndex());
        scoreFont.setItalic(true);
        maxScoreStyle.setFont(scoreFont);
        styles.put("maxScore", maxScoreStyle);

        return styles;
    }

    private static void createInfoSheet(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styles, String className) {
        XSSFSheet sheet = workbook.createSheet("Информация");

        // Настраиваем ширину колонок
        sheet.setColumnWidth(0, 4000); // Колонка A
        sheet.setColumnWidth(1, 8000); // Колонка B

        // Создаем строки
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("Учитель");
        row1.getCell(0).setCellStyle(styles.get("header"));
        Cell teacherCell = row1.createCell(1);
        teacherCell.setCellStyle(styles.get("teacher"));

        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue("Дата написания работы");
        row2.getCell(0).setCellStyle(styles.get("header"));
        Cell dateCell = row2.createCell(1);
        dateCell.setCellStyle(styles.get("teacher"));

        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue("Предмет");
        row3.getCell(0).setCellStyle(styles.get("header"));
        row3.createCell(1).setCellValue(SUBJECT);
        row3.getCell(1).setCellStyle(styles.get("normal"));

        Row row4 = sheet.createRow(3);
        row4.createCell(0).setCellValue("Класс");
        row4.getCell(0).setCellStyle(styles.get("header"));
        row4.createCell(1).setCellValue(className);
        row4.getCell(1).setCellStyle(styles.get("normal"));

        Row row5 = sheet.createRow(4);
        row5.createCell(0).setCellValue("Тип");
        row5.getCell(0).setCellStyle(styles.get("header"));
        row5.createCell(1).setCellValue("Входная работа");
        row5.getCell(1).setCellStyle(styles.get("normal"));

        // Добавляем информацию о максимальных баллах
        Row scoreInfoRow = sheet.createRow(5);
        scoreInfoRow.createCell(0).setCellValue("Макс. баллы за задания:");
        scoreInfoRow.getCell(0).setCellStyle(styles.get("header"));

        StringBuilder scoresBuilder = new StringBuilder();
        for (int i = 0; i < CURRENT_TASKS_COUNT; i++) {
            if (i > 0) scoresBuilder.append(", ");
            scoresBuilder.append(i + 1).append("=").append(getMaxScoreForTask(i));
        }
        scoreInfoRow.createCell(1).setCellValue(scoresBuilder.toString());
        scoreInfoRow.getCell(1).setCellStyle(styles.get("normal"));

        // Добавляем комментарий
        Row commentRow = sheet.createRow(6);
        commentRow.createCell(0).setCellValue("Примечание:");
        commentRow.getCell(0).setCellStyle(styles.get("header"));
        commentRow.createCell(1).setCellValue("Заполнить поля 'Учитель' и 'Дата' перед началом работы");
        commentRow.getCell(1).setCellStyle(styles.get("normal"));

        // Скрываем строки 7-20
        for (int i = 7; i < 20; i++) {
            Row row = sheet.createRow(i);
            if (row != null) {
                row.createCell(0);
                row.createCell(1);
                row.setZeroHeight(true);
            }
        }

        // Скрываем колонки C-Z
        for (int i = 2; i < 26; i++) {
            sheet.setColumnHidden(i, true);
        }

        // Авторазмер для первой колонки
        sheet.autoSizeColumn(0);
    }

    private static void createDataCollectionSheet(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styles,
                                                  String className, List<Student> students) {
        XSSFSheet sheet = workbook.createSheet("Сбор информации");

        // ===== ЗАГОЛОВКИ =====
        Row headerRow = sheet.createRow(0);
        String[] headers = {
                "№",
                "ФИО ученика",
                "Присутствие",
                "Вариант"
        };

        // Создаем основные заголовки
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Заголовок "Баллы за задания" (объединенная ячейка)
        Cell taskHeaderCell = headerRow.createCell(headers.length);
        taskHeaderCell.setCellValue("Баллы за задания");
        taskHeaderCell.setCellStyle(styles.get("header"));

        // Объединяем ячейку "Баллы за задания"
        sheet.addMergedRegion(new CellRangeAddress(0, 0, headers.length, headers.length + CURRENT_TASKS_COUNT));

        // Заголовок для итогового балла (после всех заданий)
        Cell totalCell = headerRow.createCell(headers.length + CURRENT_TASKS_COUNT + 1);
        totalCell.setCellValue("Итог");
        totalCell.setCellStyle(styles.get("header"));

        // ===== ВТОРАЯ СТРОКА: номера заданий =====
        Row taskNumberRow = sheet.createRow(1);

        // Пустые ячейки для первых 4 колонок
        for (int i = 0; i < headers.length; i++) {
            Cell emptyCell = taskNumberRow.createCell(i);
            emptyCell.setCellStyle(styles.get("header"));
        }

        // Номера заданий (1-19 или другое количество)
        int taskStartCol = headers.length; // Начинаем с колонки E (индекс 4)
        for (int i = 1; i <= CURRENT_TASKS_COUNT; i++) {
            Cell taskNumberCell = taskNumberRow.createCell(taskStartCol + i);
            taskNumberCell.setCellValue(i);
            taskNumberCell.setCellStyle(styles.get("taskHeader"));
        }

        // Пустая ячейка для итога на второй строке
        Cell emptyTotalCell = taskNumberRow.createCell(taskStartCol + CURRENT_TASKS_COUNT + 1);
        emptyTotalCell.setCellStyle(styles.get("header"));

        // ===== ТРЕТЬЯ СТРОКА: максимальные баллы за задания =====
        Row maxScoresRow = sheet.createRow(2);

        // Пустые ячейки для первых 4 колонок
        for (int i = 0; i < headers.length; i++) {
            Cell emptyCell = maxScoresRow.createCell(i);
            emptyCell.setCellStyle(styles.get("header"));
        }

        // Максимальные баллы за каждое задание
        for (int i = 1; i <= CURRENT_TASKS_COUNT; i++) {
            Cell maxScoreCell = maxScoresRow.createCell(taskStartCol + i);
            int maxScore = getMaxScoreForTask(i - 1); // i-1 потому что массив с 0
            maxScoreCell.setCellValue(maxScore);
            maxScoreCell.setCellStyle(styles.get("maxScore"));
        }

        // Пустая ячейка для итога на третьей строке
        Cell emptyTotalCell2 = maxScoresRow.createCell(taskStartCol + CURRENT_TASKS_COUNT + 1);
        emptyTotalCell2.setCellStyle(styles.get("header"));

        // ===== ДАННЫЕ УЧЕНИКОВ (начинаются с 4 строки) =====
        int studentCount = Math.min(students.size(), MAX_STUDENTS);
        int firstStudentRow = 3; // Строки с данными начинаются с индекса 3 (4-я строка в Excel)

        for (int i = 0; i < studentCount; i++) {
            Row row = sheet.createRow(firstStudentRow + i);
            Student student = students.get(i);

            // № п/п
            Cell numCell = row.createCell(0);
            numCell.setCellValue(i + 1);
            numCell.setCellStyle(styles.get("center"));

            // ФИО
            Cell fioCell = row.createCell(1);
            fioCell.setCellValue(student.getFio());
            fioCell.setCellStyle(styles.get("normal"));

            // Присутствие (выпадающий список)
            Cell presenceCell = row.createCell(2);
            presenceCell.setCellStyle(styles.get("center"));

            // Вариант (выпадающий список)
            Cell variantCell = row.createCell(3);
            variantCell.setCellStyle(styles.get("center"));

            // Ячейки для баллов за задания
            for (int taskNum = 1; taskNum <= CURRENT_TASKS_COUNT; taskNum++) {
                Cell taskCell = row.createCell(taskStartCol + taskNum);
                taskCell.setCellStyle(styles.get("center"));
                // Валидация будет настроена позже
            }

            // Итоговая ячейка (формула суммы с учетом разных весов заданий)
            Cell totalScoreCell = row.createCell(taskStartCol + CURRENT_TASKS_COUNT + 1);
            totalScoreCell.setCellStyle(styles.get("center"));

            // Формула для итога (Excel строки начинаются с 1, а POI с 0)
            int excelRowNum = firstStudentRow + i + 1; // +1 потому что Excel считает с 1
            String formula = createTotalFormula(taskStartCol, excelRowNum);
            totalScoreCell.setCellFormula(formula);
        }

        // Добавляем пустые строки до максимума (начиная с текущей позиции)
        for (int i = studentCount; i < MAX_STUDENTS; i++) {
            Row row = sheet.createRow(firstStudentRow + i);
            for (int col = 0; col < taskStartCol + CURRENT_TASKS_COUNT + 2; col++) {
                Cell cell = row.createCell(col);
                cell.setCellStyle(styles.get("normal"));
            }
            row.getCell(0).setCellValue(i + 1);
            row.getCell(0).setCellStyle(styles.get("center"));
        }

        // Скрываем неиспользуемые строки (после максимального количества учеников + заголовки)
        int lastDataRow = firstStudentRow + MAX_STUDENTS - 1;
        for (int i = lastDataRow + 1; i < 100; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            row.setZeroHeight(true);
        }

        // ===== НАСТРОЙКА ВЫПАДАЮЩИХ СПИСКОВ =====
        setupDataValidation(sheet, studentCount, firstStudentRow);

        // ===== НАСТРОЙКА УСЛОВНОГО ФОРМАТИРОВАНИЯ =====
        setupConditionalFormatting(workbook, sheet, studentCount, firstStudentRow);

        // ===== НАСТРОЙКА ШИРИНЫ КОЛОНОК =====
        sheet.setColumnWidth(0, 1000);  // №
        sheet.setColumnWidth(1, 8000);  // ФИО
        sheet.setColumnWidth(2, 3000);  // Присутствие
        sheet.setColumnWidth(3, 2500);  // Вариант

        // Ширина для заданий
        for (int i = 0; i < CURRENT_TASKS_COUNT; i++) {
            sheet.setColumnWidth(headers.length + 1 + i, 1500);
        }

        sheet.setColumnWidth(headers.length + CURRENT_TASKS_COUNT + 1, 2000); // Итог

        // Скрываем неиспользуемые колонки (включая колонки для заданий, которых нет)
        int lastUsedColumn = headers.length + CURRENT_TASKS_COUNT + 1;
        int maxPossibleColumn = headers.length + MAX_TASKS + 1; // Максимально возможная колонка

        for (int i = lastUsedColumn + 1; i <= maxPossibleColumn; i++) {
            sheet.setColumnHidden(i, true);
        }

        // Замораживаем область с заголовками (3 строки заголовков)
        sheet.createFreezePane(0, 3);
    }

    private static void setupDataValidation(XSSFSheet sheet, int studentCount, int firstStudentRow) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);

        // Список для присутствия
        String[] presenceOptions = {"Был", "Не был"};
        XSSFDataValidationConstraint presenceConstraint =
                (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(presenceOptions);

        // Список для вариантов
        List<String> variantOptions = new ArrayList<>();
        for (int i = 1; i <= VARIANTS_COUNT; i++) {
            variantOptions.add("Вариант " + i);
        }
        XSSFDataValidationConstraint variantConstraint =
                (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(variantOptions.toArray(new String[0]));

        // Применяем валидацию к столбцу присутствия (колонка C)
        CellRangeAddressList presenceRange = new CellRangeAddressList(
                firstStudentRow, firstStudentRow + studentCount - 1, 2, 2);
        XSSFDataValidation presenceValidation =
                (XSSFDataValidation) dvHelper.createValidation(presenceConstraint, presenceRange);
        presenceValidation.setShowErrorBox(true);
        presenceValidation.setEmptyCellAllowed(true);
        sheet.addValidationData(presenceValidation);

        // Применяем валидацию к столбцу вариантов (колонка D)
        CellRangeAddressList variantRange = new CellRangeAddressList(
                firstStudentRow, firstStudentRow + studentCount - 1, 3, 3);
        XSSFDataValidation variantValidation =
                (XSSFDataValidation) dvHelper.createValidation(variantConstraint, variantRange);
        variantValidation.setShowErrorBox(true);
        variantValidation.setEmptyCellAllowed(true);
        sheet.addValidationData(variantValidation);

        // Применяем валидацию к ячейкам заданий (колонки E и далее)
        int taskStartCol = 4; // Колонка E
        for (int taskNum = 1; taskNum <= CURRENT_TASKS_COUNT; taskNum++) {
            CellRangeAddressList scoreRange = new CellRangeAddressList(
                    firstStudentRow, firstStudentRow + studentCount - 1,
                    taskStartCol + taskNum, taskStartCol + taskNum);

            // Создаем выпадающий список с возможными баллами от 0 до максимального
            int maxScore = getMaxScoreForTask(taskNum - 1);
            List<String> scoreOptions = new ArrayList<>();
            for (int score = 0; score <= maxScore; score++) {
                scoreOptions.add(String.valueOf(score));
            }

            XSSFDataValidationConstraint scoreConstraint =
                    (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(scoreOptions.toArray(new String[0]));

            XSSFDataValidation scoreValidation =
                    (XSSFDataValidation) dvHelper.createValidation(scoreConstraint, scoreRange);
            scoreValidation.setShowErrorBox(true);
            scoreValidation.setEmptyCellAllowed(true);
            sheet.addValidationData(scoreValidation);
        }
    }

    private static void setupConditionalFormatting(Workbook workbook, Sheet sheet, int studentCount, int firstStudentRow) {
        // Правило для подсветки "Был" (зеленый)
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();

        // Условное форматирование для присутствия
        // Excel строки начинаются с 1, а POI с 0
        int firstExcelRow = firstStudentRow + 1;

        ConditionalFormattingRule rule1 = scf.createConditionalFormattingRule(
                "EXACT($C" + firstExcelRow + ",\"Был\")");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(GREEN_COLOR);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        ConditionalFormattingRule rule2 = scf.createConditionalFormattingRule(
                "EXACT($C" + firstExcelRow + ",\"Не был\")");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(RED_COLOR);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // Применяем к колонке присутствия
        CellRangeAddress[] ranges = {
                new CellRangeAddress(firstStudentRow, firstStudentRow + studentCount - 1, 2, 2)
        };
        scf.addConditionalFormatting(ranges, rule1, rule2);
    }

    private static int getMaxScoreForTask(int taskIndex) {
        // Если индекс выходит за пределы массива, возвращаем 1 (значение по умолчанию)
        if (taskIndex >= 0 && taskIndex < MAX_SCORES_PER_TASK.length) {
            return MAX_SCORES_PER_TASK[taskIndex];
        }
        return 1; // Значение по умолчанию для всех остальных заданий
    }

    private static String createTotalFormula(int taskStartCol, int excelRowNum) {
        StringBuilder formula = new StringBuilder("SUM(");

        // Формируем список ячеек для суммирования
        for (int i = 1; i <= CURRENT_TASKS_COUNT; i++) {
            if (i > 1) formula.append(",");
            String cellRef = CellReference.convertNumToColString(taskStartCol + i) + excelRowNum;
            formula.append(cellRef);
        }

        formula.append(")");
        return formula.toString();
    }

    static class Student {
        private final String fio;
        private final String className;

        public Student(String fio, String className) {
            this.fio = fio;
            this.className = className;
        }

        public String getFio() {
            return fio;
        }

        public String getClassName() {
            return className;
        }
    }
}