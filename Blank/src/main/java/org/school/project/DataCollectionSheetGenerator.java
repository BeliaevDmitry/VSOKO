package org.school.project;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.nio.file.*;
import java.util.*;

public class DataCollectionSheetGenerator {

    // ===== БАЗОВЫЕ ПУТИ =====
    private static final String OUTPUT_BASE_FOLDER_GENERAL =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\ВСОКО\\Работы\\{предмет}\\Формы сбора\\{класс}";
    private static final String OUTPUT_BASE_FOLDER_OGE =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\ВСОКО\\Работы\\ОГЭ\\{предмет}\\Формы сбора\\{класс}";

    // ===== НАСТРОЙКИ =====
    private static final String EXCEL_TEMPLATE_PATH =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\БД основные\\Реестр контингента.xlsx";
    private static final String ADMIN_EXCEL_PATH =
            "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ №7\\ОГЭ 2026\\ОГЭ 2026 админ.xlsx";
    private static final String PARALLEL = "9";
    private static final String SCHOOL_NAME = "ГБОУ №7";
    private static final String ACADEMIC_YEAR = "2025-2026";
    private static final int MAX_STUDENTS = 500;
    private static final int VARIANTS_COUNT = 6;

    // Для ручного режима
    private static final String MANUAL_SUBJECT = "Английский язык";
    private static final int[] MANUAL_MAX_SCORES = {
            1, 1, 1, 1, 1,
            1, 1, 1, 1, 1,
            1, 1, 1, 1, 1,
            1, 1, 1, 1
    };

    // ===== ЦВЕТА (индексы из IndexedColors) =====
    private static final short GREEN_COLOR = IndexedColors.LIGHT_GREEN.getIndex();
    private static final short RED_COLOR = IndexedColors.RED.getIndex();
    private static final short LIGHT_YELLOW = IndexedColors.LIGHT_YELLOW.getIndex();
    private static final short LIGHT_TURQUOISE = IndexedColors.TURQUOISE.getIndex();
    private static final short LIGHT_CORNFLOWER_BLUE = IndexedColors.PALE_BLUE.getIndex();
    private static final short LIGHT_GRAY = IndexedColors.GREY_25_PERCENT.getIndex();

    // ===== ДАННЫЕ ДЛЯ ОГЭ (из таблицы) =====
    private static final Map<String, int[]> OGE_SUBJECTS = new LinkedHashMap<>();

    static {
        OGE_SUBJECTS.put("Математика", new int[]{
                1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2
        });
        OGE_SUBJECTS.put("Русский язык", new int[]{
                6,1,1,1,1,1,1,1,1,1,1,1,7,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1
        });
        OGE_SUBJECTS.put("Обществознание", new int[]{
                2,1,1,1,3,2,1,1,1,1,1,4,1,1,2,1,1,1,1,1,2,2,3,2,1
        });
        OGE_SUBJECTS.put("Физика", new int[]{
                2,2,1,2,1,1,1,1,1,1,1,2,2,2,1,2,3,2,2,3,3,3,1,1,1
        });
        OGE_SUBJECTS.put("Химия", new int[]{
                1,1,1,2,1,1,1,1,2,2,1,2,1,1,1,1,2,1,1,3,3,3,5,1,1,1
        });
        OGE_SUBJECTS.put("Биология", new int[]{
                1,1,1,2,2,1,2,1,2,2,2,1,3,1,1,2,2,2,2,1,2,2,2,3,3,3,3
        });
        OGE_SUBJECTS.put("История", new int[]{
                2,1,1,2,1,1,2,1,1,1,1,1,2,1,1,1,2,2,2,2,2,3,3,1,1,1,1,1,1,1,1,1,1,1,1,1,1
        });
        OGE_SUBJECTS.put("География", new int[]{
                1,1,1,1,1,1,1,1,1,1,1,2,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1
        });
        OGE_SUBJECTS.put("Информатика", new int[]{
                1,1,1,1,1,1,1,1,1,1,1,1,2,3,2,2
        });
        OGE_SUBJECTS.put("Литература", new int[]{
                5,5,5,8,17
        });
        OGE_SUBJECTS.put("Английский язык", new int[]{
                1,1,1,1,5,1,1,1,1,1,1,6,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,10,2,6,7
        });
    }

    public static void main(String[] args) {
        try {
            System.out.println("Выберите режим работы:");
            System.out.println("1 - ОГЭ (автоматические баллы по предмету)");
            System.out.println("2 - ЕГЭ (заглушка)");
            System.out.println("3 - Ручной ввод (используются константы в коде)");
            System.out.print("Ваш выбор: ");
            Scanner scanner = new Scanner(System.in);
            int mode = scanner.nextInt();
            scanner.nextLine();

            if (mode == 2) {
                System.out.println("Режим ЕГЭ находится в разработке. Программа завершена.");
                return;
            }

            String subjectName;
            int[] maxScores;
            int tasksCount;
            String basePathTemplate;
            Map<String, Set<String>> ogeStudentSubjects = null;

            if (mode == 1) {
                System.out.println("\nДоступные предметы ОГЭ:");
                List<String> subjects = new ArrayList<>(OGE_SUBJECTS.keySet());
                for (int i = 0; i < subjects.size(); i++) {
                    System.out.println((i + 1) + " - " + subjects.get(i));
                }
                System.out.print("Выберите номер предмета: ");
                int choice = scanner.nextInt();
                scanner.nextLine();
                if (choice < 1 || choice > subjects.size()) {
                    System.out.println("Неверный выбор. Завершение.");
                    return;
                }
                subjectName = subjects.get(choice - 1);
                maxScores = OGE_SUBJECTS.get(subjectName);
                tasksCount = maxScores.length;
                basePathTemplate = OUTPUT_BASE_FOLDER_OGE;

                ogeStudentSubjects = loadOGEStudentsMap();
                System.out.println("Загружены данные о сдающих ОГЭ: " + ogeStudentSubjects.size() + " учеников");

                System.out.println("Выбран предмет: " + subjectName);
                System.out.println("Количество заданий: " + tasksCount);
                System.out.println("Сумма баллов: " + Arrays.stream(maxScores).sum());
            } else {
                subjectName = MANUAL_SUBJECT;
                maxScores = MANUAL_MAX_SCORES;
                tasksCount = maxScores.length;
                basePathTemplate = OUTPUT_BASE_FOLDER_GENERAL;
                System.out.println("Ручной режим. Предмет: " + subjectName);
                System.out.println("Количество заданий: " + tasksCount);
            }

            // NEW: выбор типа отчёта
            System.out.println("\nВыберите тип отчёта:");
            System.out.println("1 - Отдельные файлы по каждому классу");
            System.out.println("2 - Один общий файл для всей параллели");
            System.out.print("Ваш выбор: ");
            int reportType = scanner.nextInt();
            scanner.nextLine();

            String parallelFolder = basePathTemplate
                    .replace("{предмет}", subjectName)
                    .replace("{класс}", PARALLEL + " класс");
            Files.createDirectories(Paths.get(parallelFolder));

            List<String> classes = readClassesFromExcel();
            System.out.println("Найдены классы для обработки: " + classes);

            if (reportType == 2) {
                // Общий файл для всей параллели
                createParallelDataCollectionFile(parallelFolder, subjectName, maxScores, tasksCount,
                        classes, ogeStudentSubjects);
            } else {
                // Прежний режим: по классам
                for (String className : classes) {
                    createDataCollectionFile(className, parallelFolder, subjectName, maxScores, tasksCount,
                            ogeStudentSubjects);
                }
            }

            System.out.println("\nГотово! Файлы созданы в папке: " + parallelFolder);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ==================

    private static Map<String, Set<String>> loadOGEStudentsMap() throws IOException {
        Map<String, Set<String>> studentSubjects = new HashMap<>();
        Set<String> possibleSubjects = new HashSet<>(OGE_SUBJECTS.keySet());
        possibleSubjects.add("Иностранный язык");

        try (FileInputStream fis = new FileInputStream(ADMIN_EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Выбор экзамена");
            if (sheet == null) {
                throw new IllegalArgumentException("Лист 'Выбор экзамена' не найден в файле " + ADMIN_EXCEL_PATH);
            }
            for (Row row : sheet) {
                Cell fioCell = row.getCell(1);
                Cell subjectsCell = row.getCell(3);
                if (fioCell == null || subjectsCell == null) continue;
                String fio = getCellValueAsString(fioCell).trim();
                if (fio.isEmpty()) continue;
                String subjectsStr = getCellValueAsString(subjectsCell);
                if (subjectsStr.isEmpty()) continue;

                Set<String> subjects = new HashSet<>();
                for (String subject : possibleSubjects) {
                    if (subjectsStr.contains(subject)) {
                        if (subject.equals("Иностранный язык")) {
                            subjects.add("Английский язык");
                        } else {
                            subjects.add(subject);
                        }
                    }
                }
                if (!subjects.isEmpty()) {
                    studentSubjects.put(fio, subjects);
                }
            }
        }
        return studentSubjects;
    }

    private static List<String> readClassesFromExcel() throws IOException {
        List<String> classes = new ArrayList<>();
        Set<String> uniqueClasses = new TreeSet<>();

        try (FileInputStream file = new FileInputStream(EXCEL_TEMPLATE_PATH);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Cell classCell = row.getCell(15);
                if (classCell != null) {
                    String className = getCellValueAsString(classCell).trim();
                    String normalized = normalizeClassName(className);
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
                }
                double value = cell.getNumericCellValue();
                return (value == Math.floor(value)) ? String.valueOf((int) value) : String.valueOf(value);
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
        students.sort(Comparator.comparing(Student::getFio));
        return students;
    }

    // -------------------- СОЗДАНИЕ ФАЙЛОВ ПО КЛАССАМ (старый метод) --------------------
    private static void createDataCollectionFile(String className, String parallelFolder,
                                                 String subjectName, int[] maxScores, int tasksCount,
                                                 Map<String, Set<String>> ogeStudentSubjects) throws IOException {
        List<Student> allStudentsInClass = readStudentsByClass(className);
        List<Student> filteredStudents;

        if (ogeStudentSubjects != null) {
            filteredStudents = new ArrayList<>();
            for (Student student : allStudentsInClass) {
                Set<String> subjects = ogeStudentSubjects.get(student.getFio());
                if (subjects != null && subjects.contains(subjectName)) {
                    filteredStudents.add(student);
                }
            }
            if (filteredStudents.isEmpty()) {
                System.out.println("В классе " + className + " нет учеников, сдающих " + subjectName +
                        ". Файл не создаётся.");
                return;
            }
        } else {
            filteredStudents = allStudentsInClass;
        }

        String safeClassName = className.replaceAll("[\\\\/:*?\"<>|]", "_");
        String fileName = parallelFolder + "\\Сбор_данных_" + safeClassName + "_" + subjectName + ".xlsx";

        System.out.println("\nСоздание файла для класса: " + className +
                " (учеников в классе: " + allStudentsInClass.size() +
                ", сдающих " + subjectName + ": " + filteredStudents.size() + ")");

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Map<String, XSSFCellStyle> styles = createStyles(workbook);
            createInfoSheet(workbook, styles, className, subjectName, maxScores, tasksCount);
            // addClassColumn = false для по-классовых файлов
            createDataCollectionSheet(workbook, styles, filteredStudents, maxScores, tasksCount, false);

            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }
            System.out.println("Файл создан: " + fileName);
        }
    }

    // NEW: создание ОДНОГО файла для всей параллели
    private static void createParallelDataCollectionFile(String parallelFolder, String subjectName,
                                                         int[] maxScores, int tasksCount,
                                                         List<String> allClasses,
                                                         Map<String, Set<String>> ogeStudentSubjects) throws IOException {
        List<Student> allFilteredStudents = new ArrayList<>();

        // Собираем учеников со всех классов
        for (String className : allClasses) {
            List<Student> classStudents = readStudentsByClass(className);
            if (ogeStudentSubjects != null) {
                for (Student student : classStudents) {
                    Set<String> subjects = ogeStudentSubjects.get(student.getFio());
                    if (subjects != null && subjects.contains(subjectName)) {
                        allFilteredStudents.add(student);
                    }
                }
            } else {
                allFilteredStudents.addAll(classStudents);
            }
        }

        if (allFilteredStudents.isEmpty()) {
            System.out.println("В параллели " + PARALLEL + " нет учеников, сдающих " + subjectName +
                    ". Файл не создаётся.");
            return;
        }

        // Сортируем по классу, затем по ФИО
        allFilteredStudents.sort(Comparator.comparing(Student::getClassName).thenComparing(Student::getFio));

        String fileName = parallelFolder + "\\Сбор_данных_" + PARALLEL + "_классы_" + subjectName + ".xlsx";
        System.out.println("\nСоздание общего файла для параллели " + PARALLEL +
                " (учеников: " + allFilteredStudents.size() + ")");

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Map<String, XSSFCellStyle> styles = createStyles(workbook);
            // На листе информации укажем "9 классы"
            createInfoSheet(workbook, styles, PARALLEL + " классы", subjectName, maxScores, tasksCount);
            // addClassColumn = true, чтобы добавить столбец "Класс"
            createDataCollectionSheet(workbook, styles, allFilteredStudents, maxScores, tasksCount, true);

            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }
            System.out.println("Файл создан: " + fileName);
        }
    }

    // ================== СТИЛИ ==================
    private static Map<String, XSSFCellStyle> createStyles(XSSFWorkbook workbook) {
        Map<String, XSSFCellStyle> styles = new HashMap<>();

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

        XSSFCellStyle normalStyle = workbook.createCellStyle();
        normalStyle.setBorderTop(BorderStyle.THIN);
        normalStyle.setBorderBottom(BorderStyle.THIN);
        normalStyle.setBorderLeft(BorderStyle.THIN);
        normalStyle.setBorderRight(BorderStyle.THIN);
        normalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        styles.put("normal", normalStyle);

        XSSFCellStyle centerStyle = workbook.createCellStyle();
        centerStyle.cloneStyleFrom(normalStyle);
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        styles.put("center", centerStyle);

        XSSFCellStyle teacherStyle = workbook.createCellStyle();
        teacherStyle.cloneStyleFrom(normalStyle);
        teacherStyle.setFillForegroundColor(LIGHT_YELLOW);
        teacherStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("teacher", teacherStyle);

        XSSFCellStyle greenStyle = workbook.createCellStyle();
        greenStyle.cloneStyleFrom(centerStyle);
        greenStyle.setFillForegroundColor(GREEN_COLOR);
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("green", greenStyle);

        XSSFCellStyle redStyle = workbook.createCellStyle();
        redStyle.cloneStyleFrom(centerStyle);
        redStyle.setFillForegroundColor(RED_COLOR);
        redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font redFont = workbook.createFont();
        redFont.setColor(IndexedColors.WHITE.getIndex());
        redStyle.setFont(redFont);
        styles.put("red", redStyle);

        XSSFCellStyle taskHeaderStyle = workbook.createCellStyle();
        taskHeaderStyle.cloneStyleFrom(centerStyle);
        taskHeaderStyle.setFillForegroundColor(LIGHT_CORNFLOWER_BLUE);
        taskHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font taskFont = workbook.createFont();
        taskFont.setBold(true);
        taskHeaderStyle.setFont(taskFont);
        styles.put("taskHeader", taskHeaderStyle);

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

    private static void createInfoSheet(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styles,
                                        String className, String subjectName, int[] maxScores, int tasksCount) {
        XSSFSheet sheet = workbook.createSheet("Информация");
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 8000);

        int rowNum = 0;
        Row row1 = sheet.createRow(rowNum++);
        row1.createCell(0).setCellValue("Учитель");
        row1.getCell(0).setCellStyle(styles.get("header"));
        row1.createCell(1).setCellStyle(styles.get("teacher"));

        Row row2 = sheet.createRow(rowNum++);
        row2.createCell(0).setCellValue("Дата написания работы");
        row2.getCell(0).setCellStyle(styles.get("header"));
        row2.createCell(1).setCellStyle(styles.get("teacher"));

        Row row3 = sheet.createRow(rowNum++);
        row3.createCell(0).setCellValue("Предмет");
        row3.getCell(0).setCellStyle(styles.get("header"));
        row3.createCell(1).setCellValue(subjectName);
        row3.getCell(1).setCellStyle(styles.get("normal"));

        Row row4 = sheet.createRow(rowNum++);
        row4.createCell(0).setCellValue("Класс");
        row4.getCell(0).setCellStyle(styles.get("header"));
        row4.createCell(1).setCellValue(className);
        row4.getCell(1).setCellStyle(styles.get("normal"));

        Row row5 = sheet.createRow(rowNum++);
        row5.createCell(0).setCellValue("Тип");
        row5.getCell(0).setCellStyle(styles.get("header"));
        row5.createCell(1).setCellValue("Входная работа");
        row5.getCell(1).setCellStyle(styles.get("normal"));

        Row scoreInfoRow = sheet.createRow(rowNum++);
        scoreInfoRow.createCell(0).setCellValue("Макс. баллы за задания:");
        scoreInfoRow.getCell(0).setCellStyle(styles.get("header"));
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < tasksCount; i++) {
            if (i > 0) sb.append(", ");
            sb.append(i + 1).append("=").append(maxScores[i]);
        }
        scoreInfoRow.createCell(1).setCellValue(sb.toString());
        scoreInfoRow.getCell(1).setCellStyle(styles.get("normal"));

        Row commentRow = sheet.createRow(rowNum++);
        commentRow.createCell(0).setCellValue("Примечание:");
        commentRow.getCell(0).setCellStyle(styles.get("header"));
        commentRow.createCell(1).setCellValue("Заполнить поля 'Учитель' и 'Дата' перед началом работы");
        commentRow.getCell(1).setCellStyle(styles.get("normal"));

        Row schoolRow = sheet.createRow(rowNum++);
        schoolRow.createCell(0).setCellValue("Школа");
        schoolRow.getCell(0).setCellStyle(styles.get("header"));
        schoolRow.createCell(1).setCellValue(SCHOOL_NAME);
        schoolRow.getCell(1).setCellStyle(styles.get("normal"));

        Row yearRow = sheet.createRow(rowNum++);
        yearRow.createCell(0).setCellValue("Учебный год");
        yearRow.getCell(0).setCellStyle(styles.get("header"));
        yearRow.createCell(1).setCellValue(ACADEMIC_YEAR);
        yearRow.getCell(1).setCellStyle(styles.get("normal"));

        for (int i = rowNum; i < 20; i++) {
            Row row = sheet.createRow(i);
            row.createCell(0);
            row.createCell(1);
            row.setZeroHeight(true);
        }
        for (int i = 2; i < 26; i++) sheet.setColumnHidden(i, true);
        sheet.autoSizeColumn(0);
    }

    // MODIFIED: добавлен параметр addClassColumn
    private static void createDataCollectionSheet(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styles,
                                                  List<Student> students, int[] maxScores, int tasksCount,
                                                  boolean addClassColumn) {
        XSSFSheet sheet = workbook.createSheet("Сбор информации");

        // Определяем заголовки в зависимости от addClassColumn
        List<String> headers = new ArrayList<>();
        headers.add("№");
        headers.add("ФИО ученика");
        if (addClassColumn) {
            headers.add("Класс");
        }
        headers.add("Присутствие");
        headers.add("Вариант");

        int taskStartCol = headers.size(); // колонка, с которой начинаются задания

        // Строка 0: заголовки
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(styles.get("header"));
        }
        Cell taskHeaderCell = headerRow.createCell(taskStartCol);
        taskHeaderCell.setCellValue("Баллы за задания");
        taskHeaderCell.setCellStyle(styles.get("header"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, taskStartCol, taskStartCol + tasksCount - 1));
        Cell totalCell = headerRow.createCell(taskStartCol + tasksCount);
        totalCell.setCellValue("Итог");
        totalCell.setCellStyle(styles.get("header"));

        // Строка 1: номера заданий
        Row taskNumberRow = sheet.createRow(1);
        for (int i = 0; i < headers.size(); i++) {
            taskNumberRow.createCell(i).setCellStyle(styles.get("header"));
        }
        for (int i = 0; i < tasksCount; i++) {
            Cell cell = taskNumberRow.createCell(taskStartCol + i);
            cell.setCellValue(i + 1);
            cell.setCellStyle(styles.get("taskHeader"));
        }
        taskNumberRow.createCell(taskStartCol + tasksCount).setCellStyle(styles.get("header"));

        // Строка 2: максимальные баллы
        Row maxScoresRow = sheet.createRow(2);
        for (int i = 0; i < headers.size(); i++) {
            maxScoresRow.createCell(i).setCellStyle(styles.get("header"));
        }
        for (int i = 0; i < tasksCount; i++) {
            Cell cell = maxScoresRow.createCell(taskStartCol + i);
            cell.setCellValue(maxScores[i]);
            cell.setCellStyle(styles.get("maxScore"));
        }
        maxScoresRow.createCell(taskStartCol + tasksCount).setCellStyle(styles.get("header"));

        int firstStudentRow = 3;
        int studentCount = Math.min(students.size(), MAX_STUDENTS);
        for (int i = 0; i < studentCount; i++) {
            Row row = sheet.createRow(firstStudentRow + i);
            Student student = students.get(i);

            int col = 0;
            // №
            row.createCell(col).setCellValue(i + 1);
            row.getCell(col++).setCellStyle(styles.get("center"));
            // ФИО
            row.createCell(col).setCellValue(student.getFio());
            row.getCell(col++).setCellStyle(styles.get("normal"));
            // Класс (если нужно)
            if (addClassColumn) {
                row.createCell(col).setCellValue(student.getClassName());
                row.getCell(col++).setCellStyle(styles.get("center"));
            }
            // Присутствие
            row.createCell(col++).setCellStyle(styles.get("center"));
            // Вариант
            row.createCell(col++).setCellStyle(styles.get("center"));

            // Ячейки для баллов
            for (int t = 0; t < tasksCount; t++) {
                row.createCell(taskStartCol + t).setCellStyle(styles.get("center"));
            }
            // Формула итога
            Cell totalScoreCell = row.createCell(taskStartCol + tasksCount);
            totalScoreCell.setCellStyle(styles.get("center"));
            String formula = createTotalFormula(taskStartCol, firstStudentRow + i + 1, tasksCount);
            totalScoreCell.setCellFormula(formula);
        }

        // Заполнение пустых строк до MAX_STUDENTS
        for (int i = studentCount; i < MAX_STUDENTS; i++) {
            Row row = sheet.createRow(firstStudentRow + i);
            int col = 0;
            row.createCell(col++).setCellValue(i + 1);
            row.createCell(col++).setCellStyle(styles.get("normal"));
            if (addClassColumn) row.createCell(col++).setCellStyle(styles.get("center"));
            row.createCell(col++).setCellStyle(styles.get("center"));
            row.createCell(col++).setCellStyle(styles.get("center"));
            for (int t = 0; t < tasksCount; t++) {
                row.createCell(taskStartCol + t).setCellStyle(styles.get("center"));
            }
            row.createCell(taskStartCol + tasksCount).setCellStyle(styles.get("center"));
            row.getCell(0).setCellStyle(styles.get("center"));
        }

        int lastDataRow = firstStudentRow + MAX_STUDENTS - 1;
        for (int i = lastDataRow + 1; i < 100; i++) {
            Row row = sheet.getRow(i);
            if (row == null) row = sheet.createRow(i);
            row.setZeroHeight(true);
        }

        // Проверки данных с учётом сдвига колонок
        setupDataValidation(sheet, studentCount, firstStudentRow, tasksCount, maxScores, addClassColumn);
        setupConditionalFormatting(workbook, sheet, studentCount, firstStudentRow, addClassColumn);

        // Ширина колонок
        sheet.setColumnWidth(0, 1000);
        sheet.setColumnWidth(1, 8000);
        int nextCol = 2;
        if (addClassColumn) {
            sheet.setColumnWidth(nextCol++, 3000); // Класс
        }
        sheet.setColumnWidth(nextCol++, 3000); // Присутствие
        sheet.setColumnWidth(nextCol++, 2500); // Вариант
        for (int i = 0; i < tasksCount; i++) {
            sheet.setColumnWidth(taskStartCol + i, 1500);
        }
        sheet.setColumnWidth(taskStartCol + tasksCount, 2000);

        int lastUsedColumn = taskStartCol + tasksCount;
        for (int i = lastUsedColumn + 1; i <= lastUsedColumn + 50; i++) {
            sheet.setColumnHidden(i, true);
        }
        sheet.createFreezePane(0, 3);
    }

    // MODIFIED: учёт addClassColumn для смещения индексов
    private static void setupDataValidation(XSSFSheet sheet, int studentCount, int firstStudentRow,
                                            int tasksCount, int[] maxScores, boolean addClassColumn) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);

        // Индексы колонок для проверок
        int presenceCol = addClassColumn ? 3 : 2;
        int variantCol = presenceCol + 1;

        String[] presenceOptions = {"Был", "Не был"};
        XSSFDataValidationConstraint presenceConstraint =
                (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(presenceOptions);
        CellRangeAddressList presenceRange = new CellRangeAddressList(
                firstStudentRow, firstStudentRow + studentCount - 1, presenceCol, presenceCol);
        XSSFDataValidation presenceValidation = (XSSFDataValidation) dvHelper.createValidation(presenceConstraint, presenceRange);
        presenceValidation.setShowErrorBox(true);
        presenceValidation.setEmptyCellAllowed(true);
        sheet.addValidationData(presenceValidation);

        List<String> variantOptions = new ArrayList<>();
        for (int i = 1; i <= VARIANTS_COUNT; i++) variantOptions.add("Вариант " + i);
        XSSFDataValidationConstraint variantConstraint =
                (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(variantOptions.toArray(new String[0]));
        CellRangeAddressList variantRange = new CellRangeAddressList(
                firstStudentRow, firstStudentRow + studentCount - 1, variantCol, variantCol);
        XSSFDataValidation variantValidation = (XSSFDataValidation) dvHelper.createValidation(variantConstraint, variantRange);
        variantValidation.setShowErrorBox(true);
        variantValidation.setEmptyCellAllowed(true);
        sheet.addValidationData(variantValidation);

        // Колонка начала заданий
        int taskStartCol = variantCol + 1;
        for (int t = 0; t < tasksCount; t++) {
            int max = maxScores[t];
            List<String> scores = new ArrayList<>();
            for (int s = 0; s <= max; s++) scores.add(String.valueOf(s));
            XSSFDataValidationConstraint scoreConstraint =
                    (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(scores.toArray(new String[0]));
            CellRangeAddressList scoreRange = new CellRangeAddressList(
                    firstStudentRow, firstStudentRow + studentCount - 1, taskStartCol + t, taskStartCol + t);
            XSSFDataValidation scoreValidation = (XSSFDataValidation) dvHelper.createValidation(scoreConstraint, scoreRange);
            scoreValidation.setShowErrorBox(true);
            scoreValidation.setEmptyCellAllowed(true);
            sheet.addValidationData(scoreValidation);
        }
    }

    // MODIFIED: условное форматирование для колонки "Присутствие"
    private static void setupConditionalFormatting(Workbook workbook, Sheet sheet, int studentCount,
                                                   int firstStudentRow, boolean addClassColumn) {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        int firstExcelRow = firstStudentRow + 1;

        // Колонка "Присутствие" теперь имеет индекс 2 или 3
        int presenceCol = addClassColumn ? 3 : 2;
        String colLetter = CellReference.convertNumToColString(presenceCol);

        ConditionalFormattingRule rule1 = scf.createConditionalFormattingRule(
                "EXACT($" + colLetter + firstExcelRow + ",\"Был\")");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(GREEN_COLOR);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        ConditionalFormattingRule rule2 = scf.createConditionalFormattingRule(
                "EXACT($" + colLetter + firstExcelRow + ",\"Не был\")");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(RED_COLOR);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] ranges = {
                new CellRangeAddress(firstStudentRow, firstStudentRow + studentCount - 1, presenceCol, presenceCol)
        };
        scf.addConditionalFormatting(ranges, rule1, rule2);
    }

    private static String createTotalFormula(int taskStartCol, int excelRowNum, int tasksCount) {
        StringBuilder formula = new StringBuilder("SUM(");
        for (int i = 0; i < tasksCount; i++) {
            if (i > 0) formula.append(",");
            formula.append(CellReference.convertNumToColString(taskStartCol + i)).append(excelRowNum);
        }
        formula.append(")");
        return formula.toString();
    }

    // ================== ВНУТРЕННИЙ КЛАСС СТУДЕНТА ==================
    static class Student {
        private final String fio;
        private final String className;
        public Student(String fio, String className) { this.fio = fio; this.className = className; }
        public String getFio() { return fio; }
        public String getClassName() { return className; }
    }
}