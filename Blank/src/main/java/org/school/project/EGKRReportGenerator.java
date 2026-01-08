package org.school.project;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EGKRReportGenerator {

    // ========== НАСТРОЙКИ ==========

    // Выбор типа отчета (изменяйте эту переменную)
    private static final int REPORT_TYPE = 3;
    // 1 - по основному учителю и классу,
    // 2 - по учителю практикума и группе,
    // 3 - по основному учителю и адресу

    // Выбор школы
    private static final int SCHOOL_CHOICE = 2; // 1 - Школа 1811, 2 - ГБОУ №7

    // Константы для отчета
    private static final String ACADEMIC_YEAR = "2025-2026";
    private static final String WORK_TYPE = "ЕГКР декабрь";

    // Конфигурация для школ
    private static class SchoolConfig {
        String name;
        String egkrFolder;
        String teacherDataFile;
        String teacherDataSheet;
        String outputFolder;
        String teacherColumn;        // колонка с учителем (зависит от типа отчета)
        String groupingColumn;       // колонка для группировки (класс, группа, адрес)
        String SCHOOL_NAME;

        public SchoolConfig(String name, String egkrFolder, String teacherDataFile,
                            String teacherDataSheet, String outputFolder,
                            String teacherColumn, String groupingColumn, String SCHOOL_NAME) {
            this.name = name;
            this.egkrFolder = egkrFolder;
            this.teacherDataFile = teacherDataFile;
            this.teacherDataSheet = teacherDataSheet;
            this.outputFolder = outputFolder;
            this.teacherColumn = teacherColumn;
            this.groupingColumn = groupingColumn;
            this.SCHOOL_NAME = SCHOOL_NAME;
        }
    }

    // Метод для получения конфигурации в зависимости от типа отчета
    private static SchoolConfig getSchoolConfig() {
        switch (SCHOOL_CHOICE) {
            case 1: // Школа 1811
                switch (REPORT_TYPE) {
                    case 1: // Основной учитель + класс
                        return new SchoolConfig(
                                "Школа 1811",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГЭ 2026.xlsx",
                                "Свод вертикальный",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты\\по_учителю_и_классу",
                                "Основной учитель",  // колонка с основным учителем
                                "Номер и буква класса", // группировка по классу
                                "ГБОУ №1811"
                        );
                    case 2: // Учитель практикума + группа
                        return new SchoolConfig(
                                "Школа 1811",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГЭ 2026.xlsx",
                                "Свод вертикальный",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты\\по_учителю_практикума_и_группе",
                                "Учитель практикума", // колонка с учителем практикума
                                "группа практикума",   // группировка по группе
                                "ГБОУ №1811"
                        );
                    case 3: // Основной учитель + адрес
                        return new SchoolConfig(
                                "Школа 1811",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГЭ 2026.xlsx",
                                "Свод вертикальный",
                                "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты\\по_учителю_и_адресу",
                                "Основной учитель",  // колонка с основным учителем
                                "Адрес",             // группировка по адресу
                                "ГБОУ №1811"
                        );
                    default:
                        return getDefaultConfig();
                }
            case 2: // ГБОУ №7
                switch (REPORT_TYPE) {
                    case 1: // Основной учитель + класс
                        return new SchoolConfig(
                                "ГБОУ №7",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГЭ 2026 админ.xlsx",
                                "СВОД",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты\\по_учителю_и_классу",
                                "Основной учитель",
                                "Номер и буква класса",
                                "ГБОУ №7"
                        );
                    case 2: // Учитель практикума + группа
                        return new SchoolConfig(
                                "ГБОУ №7",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГЭ 2026 админ.xlsx",
                                "СВОД",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты\\по_учителю_практикума_и_группе",
                                "Учитель практикума",
                                "группа практикума",
                                "ГБОУ №7"
                        );
                    case 3: // Основной учитель + адрес
                        return new SchoolConfig(
                                "ГБОУ №7",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГЭ 2026 админ.xlsx",
                                "СВОД",
                                "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты\\по_учителю_и_адресу",
                                "Основной учитель",
                                "Корпус",  // Для ГБОУ №7 используем "Корпус"
                                "ГБОУ №7"
                        );
                    default:
                        return getDefaultConfig();
                }
            default:
                return getDefaultConfig();
        }
    }

    private static SchoolConfig getDefaultConfig() {
        // Конфигурация по умолчанию (тип 1)
        if (SCHOOL_CHOICE == 1) {
            return new SchoolConfig(
                    "Школа 1811",
                    "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР",
                    "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГЭ 2026.xlsx",
                    "Свод вертикальный",
                    "C:\\Users\\dimah\\Yandex.Disk\\1811\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты",
                    "Основной учитель",
                    "Номер и буква класса",
                    "ГБОУ №1811"
            );
        } else {
            return new SchoolConfig(
                    "ГБОУ №7",
                    "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР",
                    "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГЭ 2026 админ.xlsx",
                    "СВОД",
                    "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ЕГЭ 2026\\ЕГКР\\ЕГКРотчёты",
                    "Основной учитель",
                    "Номер и буква класса",
                    "ГБОУ №7"
            );
        }
    }

    // Цвета для подсветки
    private static final short GREEN_COLOR = IndexedColors.LIGHT_GREEN.getIndex();
    private static final short RED_COLOR = IndexedColors.RED.getIndex();
    private static final short LIGHT_YELLOW = IndexedColors.LIGHT_YELLOW.getIndex();
    private static final short LIGHT_TURQUOISE = IndexedColors.TURQUOISE.getIndex();
    private static final short LIGHT_CORNFLOWER_BLUE = IndexedColors.PALE_BLUE.getIndex();
    private static final short LIGHT_GRAY = IndexedColors.GREY_25_PERCENT.getIndex();

    public static void main(String[] args) {
        try {
            // Получаем конфигурацию выбранной школы и типа отчета
            SchoolConfig config = getSchoolConfig();

            System.out.println("=========================================");
            System.out.println("Генератор отчетов ЕГКР (версия 2)");
            System.out.println("Школа: " + config.SCHOOL_NAME);
            System.out.println("Учебный год: " + ACADEMIC_YEAR);
            System.out.println("Тип работы: " + WORK_TYPE);
            System.out.println("Тип отчета: " + REPORT_TYPE + " - " + getReportTypeDescription());
            System.out.println("=========================================");
            System.out.println("Папка с ЕГКР: " + config.egkrFolder);
            System.out.println("Файл учителей: " + config.teacherDataFile);
            System.out.println("Лист учителей: " + config.teacherDataSheet);
            System.out.println("Колонка учителя: " + config.teacherColumn);
            System.out.println("Колонка группировки: " + config.groupingColumn);
            System.out.println("Выходная папка: " + config.outputFolder);
            System.out.println("=========================================\n");

            // Создаем папку для отчетов
            Files.createDirectories(Paths.get(config.outputFolder));

            // 1. Загружаем данные об учителях и группах
            Map<String, Map<String, String>> studentDataMap = loadTeacherData(config);

            if (studentDataMap.isEmpty()) {
                System.out.println("ВНИМАНИЕ: Не загружены данные об учителях!");
                System.out.println("Продолжить без привязки к учителям? (y/n): ");

                Scanner scanner = new Scanner(System.in);
                String answer = scanner.nextLine().trim().toLowerCase();
                if (!answer.equals("y") && !answer.equals("yes")) {
                    System.out.println("Программа остановлена.");
                    return;
                }
            }

            // 2. Обрабатываем все файлы ЕГКР
            List<EGKRData> allData = processAllEGKRFiles(config);

            if (allData.isEmpty()) {
                System.out.println("Не найдены данные ЕГКР в папке: " + config.egkrFolder);
                return;
            }

            // 3. Группируем данные по учителям и группам
            Map<String, Map<String, List<EGKRData>>> dataByTeacherAndGroup =
                    groupDataByTeacherAndGroup(allData, studentDataMap);

            // 4. Создаем отчеты для каждого учителя и группы
            int reportCount = 0;
            for (Map.Entry<String, Map<String, List<EGKRData>>> teacherEntry :
                    dataByTeacherAndGroup.entrySet()) {

                String teacherName = teacherEntry.getKey();

                for (Map.Entry<String, List<EGKRData>> groupEntry :
                        teacherEntry.getValue().entrySet()) {

                    String groupName = groupEntry.getKey();
                    List<EGKRData> groupData = groupEntry.getValue();

                    // Группируем данные по предметам
                    Map<String, List<EGKRData>> groupedBySubject =
                            groupBySubject(groupData);

                    // Создаем отдельный отчет для каждого предмета
                    for (Map.Entry<String, List<EGKRData>> subjectEntry :
                            groupedBySubject.entrySet()) {

                        String subject = subjectEntry.getKey();
                        List<EGKRData> subjectResults = subjectEntry.getValue();

                        if (subjectResults.size() > 0) {
                            createTeacherReport(config, teacherName, groupName,
                                    subject, subjectResults);
                            reportCount++;
                        }
                    }
                }
            }

            System.out.println("\n=========================================");
            System.out.println("ГОТОВО! Создано " + reportCount + " отчетов");
            System.out.println("Школа: " + config.SCHOOL_NAME);
            System.out.println("Тип отчета: " + getReportTypeDescription());
            System.out.println("Учебный год: " + ACADEMIC_YEAR);
            System.out.println("Тип работы: " + WORK_TYPE);
            System.out.println("Отчеты сохранены в папке: " + config.outputFolder);
            System.out.println("=========================================");

        } catch (Exception e) {
            System.err.println("КРИТИЧЕСКАЯ ОШИБКА: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // ========== ОСНОВНЫЕ МЕТОДЫ ==========

    private static Map<String, Map<String, String>> loadTeacherData(SchoolConfig config) throws IOException {
        Map<String, Map<String, String>> studentDataMap = new HashMap<>();

        File teacherFile = new File(config.teacherDataFile);
        if (!teacherFile.exists()) {
            System.err.println("Файл не найден: " + config.teacherDataFile);
            return studentDataMap;
        }

        try (FileInputStream file = new FileInputStream(teacherFile);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheet(config.teacherDataSheet);
            if (sheet == null) {
                // Поиск листа по частичному совпадению
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    String sheetName = workbook.getSheetName(i);
                    if (sheetName.toLowerCase().contains(config.teacherDataSheet.toLowerCase())) {
                        sheet = workbook.getSheetAt(i);
                        System.out.println("Найден похожий лист: '" + sheetName + "'");
                        break;
                    }
                }
            }

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            int headerRowNum = findHeaderRow(sheet);

            if (headerRowNum == -1) {
                System.out.println("Не найдена строка с заголовками в файле учителей");
                return studentDataMap;
            }

            Row headerRow = sheet.getRow(headerRowNum);
            Map<String, Integer> columnIndexes = new HashMap<>();

            for (Cell cell : headerRow) {
                if (cell != null) {
                    String header = getCellValue(cell, evaluator);
                    if (!header.isEmpty()) {
                        columnIndexes.put(header, cell.getColumnIndex());
                        String normalized = header.toLowerCase().replaceAll("\\s+", "");
                        columnIndexes.put(normalized, cell.getColumnIndex());
                    }
                }
            }

            System.out.println("\n=== АНАЛИЗ СТРУКТУРЫ ФАЙЛА УЧИТЕЛЕЙ ===");
            System.out.println("Найдены колонки:");
            for (Map.Entry<String, Integer> entry : columnIndexes.entrySet()) {
                if (!entry.getKey().toLowerCase().replaceAll("\\s+", "").equals(entry.getKey())) {
                    System.out.println("  [" + entry.getValue() + "] " + entry.getKey());
                }
            }

            // Ищем нужные колонки
            Integer fioIndex = findColumnIndex(columnIndexes, new String[]{"ФИО"});
            Integer classIndex = findColumnIndex(columnIndexes, new String[]{"Номер и буква класса"});
            Integer subjectIndex = findColumnIndex(columnIndexes, new String[]{"Предмет"});
            Integer teacherIndex = findColumnIndex(columnIndexes, new String[]{config.teacherColumn});
            Integer groupIndex = findColumnIndex(columnIndexes, new String[]{config.groupingColumn});

            System.out.println("\nНАЙДЕННЫЕ КОЛОНКИ:");
            System.out.println("  ФИО: " + (fioIndex != null ? "колонка " + fioIndex : "НЕ НАЙДЕНА"));
            System.out.println("  Класс: " + (classIndex != null ? "колонка " + classIndex : "НЕ НАЙДЕНА"));
            System.out.println("  Предмет: " + (subjectIndex != null ? "колонка " + subjectIndex : "НЕ НАЙДЕНА"));
            System.out.println("  Учитель (" + config.teacherColumn + "): " +
                    (teacherIndex != null ? "колонка " + teacherIndex : "НЕ НАЙДЕНА"));
            System.out.println("  Группировка (" + config.groupingColumn + "): " +
                    (groupIndex != null ? "колонка " + groupIndex : "НЕ НАЙДЕНА"));

            if (fioIndex == null || teacherIndex == null) {
                System.out.println("\nВНИМАНИЕ: Не найдены обязательные колонки!");
                return studentDataMap;
            }

            // Читаем данные
            int rowCount = 0;
            int loadedCount = 0;

            for (int i = headerRowNum + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                rowCount++;

                String fio = getCellValue(row.getCell(fioIndex), evaluator);
                String className = classIndex != null ? getCellValue(row.getCell(classIndex), evaluator) : "";
                String subject = subjectIndex != null ? getCellValue(row.getCell(subjectIndex), evaluator) : "";
                String teacher = getCellValue(row.getCell(teacherIndex), evaluator);
                String group = groupIndex != null ? getCellValue(row.getCell(groupIndex), evaluator) : "";

                // Проверяем валидность данных
                if (isValidData(fio, teacher)) {
                    // Нормализуем данные
                    String normalizedFio = normalizeFIO(fio);
                    String normalizedClass = normalizeClassName(className);
                    String normalizedSubject = normalizeSubject(subject);
                    String normalizedGroup = normalizeGroup(group, config);

                    // Создаем ключи
                    String keyWithSubject = normalizedFio + "|" + normalizedClass + "|" + normalizedSubject;

                    // Сохраняем учителя и группу
                    Map<String, String> teacherGroup = new HashMap<>();
                    teacherGroup.put("teacher", teacher);
                    teacherGroup.put("group", normalizedGroup);

                    studentDataMap.put(keyWithSubject, teacherGroup);

                    loadedCount++;

                    if (loadedCount <= 3) {
                        System.out.println("  Пример: " + fio + " (" + className +
                                ", " + subject + ") -> Учитель: " + teacher +
                                ", Группа: " + group + " -> " + normalizedGroup);
                    }
                }
            }

            System.out.println("\nИтого: обработано " + rowCount + " строк, загружено " +
                    loadedCount + " связок ученик-учитель-группа");

        } catch (Exception e) {
            System.err.println("Ошибка при загрузке данных учителей: " + e.getMessage());
            e.printStackTrace();
        }

        return studentDataMap;
    }

    private static Map<String, Map<String, List<EGKRData>>> groupDataByTeacherAndGroup(
            List<EGKRData> allData,
            Map<String, Map<String, String>> studentDataMap) {

        // Теперь группируем по учителю И группе
        // Map<Учитель, Map<Группа, List<Данные>>>
        Map<String, Map<String, List<EGKRData>>> dataByTeacherAndGroup = new HashMap<>();
        List<EGKRData> noTeacherData = new ArrayList<>();

        System.out.println("\n=== ГРУППИРОВКА ДАННЫХ ===");
        System.out.println("Тип отчета: " + REPORT_TYPE +
                " (" + getReportTypeDescription() + ")");

        int totalProcessed = 0;
        int matched = 0;
        int notMatched = 0;

        for (EGKRData data : allData) {
            totalProcessed++;
            String normalizedFio = normalizeFIO(data.getFullName());
            String normalizedClass = normalizeClassName(data.getClassName());
            String normalizedSubject = normalizeSubject(data.getSubject());

            String fullKey = normalizedFio + "|" + normalizedClass + "|" + normalizedSubject;

            Map<String, String> teacherGroupData = studentDataMap.get(fullKey);

            if (teacherGroupData != null && !teacherGroupData.isEmpty()) {
                String teacher = teacherGroupData.get("teacher");
                String group = teacherGroupData.get("group");

                if (teacher != null && !teacher.isEmpty() &&
                        !teacher.equalsIgnoreCase("нет") &&
                        !teacher.equalsIgnoreCase("не указан")) {

                    data.setTeacherName(teacher);
                    matched++;

                    // Добавляем в структуру: учитель -> группа -> данные
                    dataByTeacherAndGroup
                            .computeIfAbsent(teacher, k -> new HashMap<>())
                            .computeIfAbsent(group, k -> new ArrayList<>())
                            .add(data);

                } else {
                    noTeacherData.add(data);
                    notMatched++;
                }
            } else {
                noTeacherData.add(data);
                notMatched++;
            }
        }

        System.out.println("\n=== СТАТИСТИКА ===");
        System.out.println("Всего записей: " + allData.size());
        System.out.println("Сопоставлено: " + matched);
        System.out.println("Не сопоставлено: " + notMatched);

        if (!noTeacherData.isEmpty()) {
            System.out.println("\nНайдено " + noTeacherData.size() + " записей без привязки");
            dataByTeacherAndGroup
                    .computeIfAbsent("Без учителя", k -> new HashMap<>())
                    .computeIfAbsent("Без группы", k -> new ArrayList<>())
                    .addAll(noTeacherData);
        }

        System.out.println("\n=== ИТОГОВАЯ ГРУППИРОВКА ===");
        for (Map.Entry<String, Map<String, List<EGKRData>>> teacherEntry : dataByTeacherAndGroup.entrySet()) {
            System.out.println("Учитель: " + teacherEntry.getKey());
            for (Map.Entry<String, List<EGKRData>> groupEntry : teacherEntry.getValue().entrySet()) {
                System.out.println("  Группа '" + groupEntry.getKey() + "': " +
                        groupEntry.getValue().size() + " записей");
            }
        }

        return dataByTeacherAndGroup;
    }

    // ========== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ==========

    private static String getReportTypeDescription() {
        switch (REPORT_TYPE) {
            case 1:
                return "По основному учителю и классу";
            case 2:
                return "По учителю практикума и группе";
            case 3:
                return "По основному учителю и адресу";
            default:
                return "Неизвестный тип";
        }
    }

    private static String getGroupLabel() {
        switch (REPORT_TYPE) {
            case 1:
                return "Класс";
            case 2:
                return "Группа практикума";
            case 3:
                return "Адрес/Корпус";
            default:
                return "Группа";
        }
    }

    private static String normalizeGroup(String group, SchoolConfig config) {
        if (group == null || group.trim().isEmpty()) {
            // Если группа пустая, используем значение по умолчанию
            switch (REPORT_TYPE) {
                case 1:
                    return "Без класса";
                case 2:
                    return "Без группы";
                case 3:
                    return "Без адреса";
                default:
                    return "Без группы";
            }
        }

        return group.trim();
    }

    // ========== ОСТАЛЬНЫЕ МЕТОДЫ (нужно скопировать из оригинального класса) ==========

    // Скопируйте сюда ВСЕ остальные методы из оригинального класса EGKRReportGenerator:
    // - getCellValue, isValidData, findColumnIndex, findHeaderRow
    // - processAllEGKRFiles, processEGKRFile, extractSubject, extractSubjectFromFileName
    // - extractDateFromLine, normalizeFIO, normalizeClassName, normalizeSubject
    // - analyzeMaxScores, parseTaskScores, createStyles, createInfoSheet, createDataCollectionSheet
    // - groupBySubject, createTeacherReport
    // - класс EGKRData

    private static void createTeacherReport(SchoolConfig config, String teacherName,
                                            String groupName, String subject,
                                            List<EGKRData> results) throws IOException {
        if (results.isEmpty()) {
            System.out.println("Нет данных для создания отчета: " + teacherName +
                    " - " + groupName + " - " + subject);
            return;
        }

        // Создаем безопасное имя файла
        String safeTeacherName = teacherName.replaceAll("[\\\\/:*?\"<>|]", "_");
        String safeGroupName = groupName.replaceAll("[\\\\/:*?\"<>|]", "_");
        String safeSubject = subject.replaceAll("[\\\\/:*?\"<>|]", "_");

        String fileName = config.outputFolder + "\\" +
                safeTeacherName + "_" + safeSubject + "_" + safeGroupName + ".xlsx";

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Map<String, XSSFCellStyle> styles = createStyles(workbook);

            // ===== ЛИСТ 1: Информация =====
            createInfoSheet(workbook, styles, teacherName, subject, groupName, results, config);

            // ===== ЛИСТ 2: Сбор информации =====
            createDataCollectionSheet(workbook, styles, subject, groupName, results);

            // Сохраняем файл
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            }

            System.out.println("Создан отчет: " + fileName +
                    " (" + results.size() + " учеников, группа: " + groupName + ")");
        } catch (Exception e) {
            System.err.println("Ошибка при создании отчета " + fileName + ": " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void createDataCollectionSheet(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styles,
                                                  String subject, String className, List<EGKRData> results) {
        XSSFSheet sheet = workbook.createSheet("Сбор информации");

        // Анализируем структуру заданий
        List<Integer> maxScores = results.isEmpty() ? new ArrayList<>() : analyzeMaxScores(results.get(0));
        int taskCount = maxScores.size();

        if (taskCount == 0) {
            // Если не удалось определить количество заданий, используем разумное значение
            taskCount = Math.max(20, results.size() > 0 ? (int) results.get(0).getPrimaryScore() : 20);
            maxScores = new ArrayList<>();
            for (int i = 0; i < taskCount; i++) {
                maxScores.add(1);
            }
        }

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

        // Заголовок "Баллы за задания"
        Cell taskHeaderCell = headerRow.createCell(headers.length);
        taskHeaderCell.setCellValue("Баллы за задания");
        taskHeaderCell.setCellStyle(styles.get("header"));

        // Объединяем ячейку "Баллы за задания"
        if (taskCount > 0) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, headers.length, headers.length + taskCount - 1));
        }

        // Заголовок для итогового балла
        Cell totalCell = headerRow.createCell(headers.length + taskCount);
        totalCell.setCellValue("Итог");
        totalCell.setCellStyle(styles.get("header"));

        // ===== ВТОРАЯ СТРОКА: номера заданий =====
        Row taskNumberRow = sheet.createRow(1);

        // Пустые ячейки для первых 4 колонок
        for (int i = 0; i < headers.length; i++) {
            taskNumberRow.createCell(i).setCellStyle(styles.get("header"));
        }

        // Номера заданий
        for (int i = 0; i < taskCount; i++) {
            Cell taskNumberCell = taskNumberRow.createCell(headers.length + i);
            taskNumberCell.setCellValue(i + 1);
            taskNumberCell.setCellStyle(styles.get("taskHeader"));
        }

        // Пустая ячейка для итога
        taskNumberRow.createCell(headers.length + taskCount).setCellStyle(styles.get("header"));

        // ===== ТРЕТЬЯ СТРОКА: максимальные баллы за задания =====
        Row maxScoresRow = sheet.createRow(2);

        // Пустые ячейки для первые 4 колонок
        for (int i = 0; i < headers.length; i++) {
            maxScoresRow.createCell(i).setCellStyle(styles.get("header"));
        }

        // Максимальные баллы за каждое задание
        for (int i = 0; i < taskCount; i++) {
            Cell maxScoreCell = maxScoresRow.createCell(headers.length + i);
            int maxScore = (i < maxScores.size()) ? maxScores.get(i) : 1;
            maxScoreCell.setCellValue(maxScore);

            // Стиль для максимальных баллов
            XSSFCellStyle maxScoreStyle = workbook.createCellStyle();
            maxScoreStyle.cloneStyleFrom(styles.get("center"));
            maxScoreStyle.setFillForegroundColor(LIGHT_GRAY);
            maxScoreStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font scoreFont = workbook.createFont();
            scoreFont.setColor(IndexedColors.DARK_RED.getIndex());
            scoreFont.setItalic(true);
            maxScoreStyle.setFont(scoreFont);
            maxScoreCell.setCellStyle(maxScoreStyle);
        }

        // Пустая ячейка для итога
        maxScoresRow.createCell(headers.length + taskCount).setCellStyle(styles.get("header"));

        // ===== ДАННЫЕ УЧЕНИКОВ =====
        int firstStudentRow = 3;

        // Сортируем учеников по ФИО
        results.sort(Comparator.comparing(EGKRData::getFullName));

        for (int i = 0; i < results.size(); i++) {
            EGKRData data = results.get(i);
            Row row = sheet.createRow(firstStudentRow + i);

            // № п/п
            Cell numCell = row.createCell(0);
            numCell.setCellValue(i + 1);
            numCell.setCellStyle(styles.get("center"));

            // ФИО
            Cell fioCell = row.createCell(1);
            fioCell.setCellValue(data.getFullName());
            fioCell.setCellStyle(styles.get("normal"));

            // Присутствие - всегда "Был" для ЕГКР
            Cell presenceCell = row.createCell(2);
            presenceCell.setCellValue("Был");
            presenceCell.setCellStyle(styles.get("green"));

            // Вариант - оставляем пустым
            Cell variantCell = row.createCell(3);
            variantCell.setCellStyle(styles.get("center"));

            // Баллы за задания
            List<Integer> taskScores = parseTaskScores(data, maxScores);

            for (int taskNum = 0; taskNum < taskCount; taskNum++) {
                Cell taskCell = row.createCell(headers.length + taskNum);
                int score = 0;
                if (taskNum < taskScores.size()) {
                    score = taskScores.get(taskNum);
                }
                taskCell.setCellValue(score);
                taskCell.setCellStyle(styles.get("center"));
            }

            // Итоговый балл
            Cell totalScoreCell = row.createCell(headers.length + taskCount);
            totalScoreCell.setCellValue(data.getPrimaryScore());
            totalScoreCell.setCellStyle(styles.get("center"));
        }

        // ===== НАСТРОЙКА ШИРИНЫ КОЛОНОК =====
        sheet.setColumnWidth(0, 1000);  // №
        sheet.setColumnWidth(1, 8000);  // ФИО
        sheet.setColumnWidth(2, 3000);  // Присутствие
        sheet.setColumnWidth(3, 2500);  // Вариант

        // Ширина для заданий
        for (int i = 0; i < taskCount; i++) {
            sheet.setColumnWidth(headers.length + i, 1500);
        }

        sheet.setColumnWidth(headers.length + taskCount, 2000); // Итог

        // Замораживаем область с заголовками
        sheet.createFreezePane(0, 3);
    }

    private static Map<String, List<EGKRData>> groupBySubject(List<EGKRData> data) {
        Map<String, List<EGKRData>> grouped = new HashMap<>();

        for (EGKRData item : data) {
            String subject = item.getSubject();
            if (subject.isEmpty()) {
                subject = "Неизвестный предмет";
            }
            grouped.computeIfAbsent(subject, k -> new ArrayList<>()).add(item);
        }

        return grouped;
    }

    private static void createInfoSheet(XSSFWorkbook workbook, Map<String, XSSFCellStyle> styles,
                                        String teacherName, String subject, String groupName,
                                        List<EGKRData> results, SchoolConfig config) {
        XSSFSheet sheet = workbook.createSheet("Информация");

        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 8000);

        String workDate = results.isEmpty() ? "2025.12.11" : results.get(0).getDate();
        int rowNum = 0;

        // Строка 1: Учитель
        Row row1 = sheet.createRow(rowNum++);
        row1.createCell(0).setCellValue("Учитель");
        row1.getCell(0).setCellStyle(styles.get("header"));
        row1.createCell(1).setCellValue(teacherName);
        row1.getCell(1).setCellStyle(styles.get("teacher"));

        // Строка 2: Дата написания работы
        Row row2 = sheet.createRow(rowNum++);
        row2.createCell(0).setCellValue("Дата написания работы");
        row2.getCell(0).setCellStyle(styles.get("header"));
        row2.createCell(1).setCellValue(workDate);
        row2.getCell(1).setCellStyle(styles.get("teacher"));

        // Строка 3: Предмет
        Row row3 = sheet.createRow(rowNum++);
        row3.createCell(0).setCellValue("Предмет");
        row3.getCell(0).setCellStyle(styles.get("header"));
        row3.createCell(1).setCellValue(subject);
        row3.getCell(1).setCellStyle(styles.get("normal"));

        // Строка 4: Группировка (зависит от типа отчета)
        Row row4 = sheet.createRow(rowNum++);
        row4.createCell(0).setCellValue("класс");
        row4.getCell(0).setCellStyle(styles.get("header"));
        row4.createCell(1).setCellValue(groupName);
        row4.getCell(1).setCellStyle(styles.get("normal"));

        // Строка 5: Тип
        Row row5 = sheet.createRow(rowNum++);
        row5.createCell(0).setCellValue("Тип");
        row5.getCell(0).setCellStyle(styles.get("header"));
        row5.createCell(1).setCellValue(WORK_TYPE);
        row5.getCell(1).setCellStyle(styles.get("normal"));

        // Строка 6: Максимальные баллы за задания (если есть данные)
        if (!results.isEmpty()) {
            EGKRData sampleData = results.get(0);
            List<Integer> maxScores = analyzeMaxScores(sampleData);

            Row row6 = sheet.createRow(rowNum++);
            row6.createCell(0).setCellValue("Макс. баллы за задания:");
            row6.getCell(0).setCellStyle(styles.get("header"));

            StringBuilder scoresBuilder = new StringBuilder();
            for (int i = 0; i < maxScores.size(); i++) {
                if (i > 0) scoresBuilder.append(", ");
                scoresBuilder.append(i + 1).append("=").append(maxScores.get(i));
            }
            row6.createCell(1).setCellValue(scoresBuilder.toString());
            row6.getCell(1).setCellStyle(styles.get("normal"));
        }

        // Строка 7: Примечание (ВСЕГДА есть)
        Row row7 = sheet.createRow(rowNum++);
        row7.createCell(0).setCellValue("Примечание:");
        row7.getCell(0).setCellStyle(styles.get("header"));
        row7.createCell(1).setCellValue("Заполнено автоматически на основе данных ЕГКР");
        row7.getCell(1).setCellStyle(styles.get("normal"));

        // Строка 8: Школа (ВСЕГДА есть)
        Row row8 = sheet.createRow(rowNum++);
        row8.createCell(0).setCellValue("Школа:");
        row8.getCell(0).setCellStyle(styles.get("header"));
        row8.createCell(1).setCellValue(config.SCHOOL_NAME);
        row8.getCell(1).setCellStyle(styles.get("normal"));

        // Строка 9: Учебный год (ВСЕГДА есть)
        Row row9 = sheet.createRow(rowNum++);
        row9.createCell(0).setCellValue("Учебный год:");
        row9.getCell(0).setCellStyle(styles.get("header"));
        row9.createCell(1).setCellValue(ACADEMIC_YEAR);
        row9.getCell(1).setCellStyle(styles.get("normal"));
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

        // Стиль для заголовков заданий
        XSSFCellStyle taskHeaderStyle = workbook.createCellStyle();
        taskHeaderStyle.cloneStyleFrom(centerStyle);
        taskHeaderStyle.setFillForegroundColor(LIGHT_CORNFLOWER_BLUE);
        taskHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font taskFont = workbook.createFont();
        taskFont.setBold(true);
        taskHeaderStyle.setFont(taskFont);
        styles.put("taskHeader", taskHeaderStyle);

        return styles;
    }

    private static List<Integer> parseTaskScores(EGKRData data, List<Integer> maxScores) {
        List<Integer> scores = new ArrayList<>();

        // Парсим задания с кратким ответом
        if (data.getShortAnswerTasks() != null && !data.getShortAnswerTasks().isEmpty()) {
            String shortTasks = data.getShortAnswerTasks();
            for (char c : shortTasks.toCharArray()) {
                if (c == '+') {
                    scores.add(1);
                } else if (c == '-') {
                    scores.add(0);
                } else if (Character.isDigit(c)) {
                    scores.add(Character.getNumericValue(c));
                }
            }
        }

        // Парсим задания с развернутым ответом
        if (data.getLongAnswerTasks() != null && !data.getLongAnswerTasks().isEmpty()) {
            String longTasks = data.getLongAnswerTasks();
            // Пример формата: 1(2)2(2)2(2)2(2)1(3)0(3)0(2)0(3)0(3)
            Pattern pattern = Pattern.compile("(\\d+)\\((\\d+)\\)");
            Matcher matcher = pattern.matcher(longTasks);

            while (matcher.find()) {
                try {
                    int score = Integer.parseInt(matcher.group(1));
                    scores.add(score);
                } catch (NumberFormatException e) {
                    scores.add(0);
                }
            }
        }

        // Если количество оценок не совпадает с количеством заданий, дополняем нулями
        while (scores.size() < maxScores.size()) {
            scores.add(0);
        }

        return scores;
    }


    private static List<Integer> analyzeMaxScores(EGKRData data) {
        List<Integer> maxScores = new ArrayList<>();

        // Анализируем задания с кратким ответом
        if (data.getShortAnswerTasks() != null && !data.getShortAnswerTasks().isEmpty()) {
            String shortTasks = data.getShortAnswerTasks();
            for (char c : shortTasks.toCharArray()) {
                if (c == '+' || c == '-' || c == '0' || c == '1' || c == '2' || c == '3') {
                    // Для заданий с кратким ответом определяем максимальный балл по символу
                    if (c == '+' || c == '-') {
                        maxScores.add(1); // +/- обычно 1 балл
                    } else if (Character.isDigit(c)) {
                        // Цифра может означать максимальный балл или набранный балл
                        // Будем считать это максимальным баллом для задания
                        maxScores.add(Character.getNumericValue(c));
                    }
                }
            }
        }

        // Анализируем задания с развернутым ответом
        if (data.getLongAnswerTasks() != null && !data.getLongAnswerTasks().isEmpty()) {
            String longTasks = data.getLongAnswerTasks();
            // Пример формата: 1(2)2(2)2(2)2(2)1(3)0(3)0(2)0(3)0(3)
            Pattern pattern = Pattern.compile("\\((\\d+)\\)");
            Matcher matcher = pattern.matcher(longTasks);

            while (matcher.find()) {
                try {
                    int maxScore = Integer.parseInt(matcher.group(1));
                    maxScores.add(maxScore);
                } catch (NumberFormatException e) {
                    maxScores.add(1); // дефолтное значение
                }
            }
        }

        // Если не удалось определить, используем дефолтные значения
        if (maxScores.isEmpty()) {
            // Для ЕГЭ по истории обычно 25 заданий, по русскому - 27 и т.д.
            // Будем использовать 20 как разумное значение по умолчанию
            int defaultTaskCount = 20;
            for (int i = 0; i < defaultTaskCount; i++) {
                maxScores.add(1);
            }
        }

        return maxScores;
    }

    private static String extractDateFromLine(String line) {
        if (line == null || line.isEmpty()) return "";

        Pattern pattern = Pattern.compile("\\d{4}\\.\\d{2}\\.\\d{2}");
        Matcher matcher = pattern.matcher(line);

        if (matcher.find()) {
            return matcher.group();
        }

        // Пробуем другие форматы
        pattern = Pattern.compile("\\d{2}\\.\\d{2}\\.\\d{4}");
        matcher = pattern.matcher(line);

        if (matcher.find()) {
            String date = matcher.group();
            String[] parts = date.split("\\.");
            if (parts.length == 3) {
                return parts[2] + "." + parts[1] + "." + parts[0];
            }
        }

        return "";
    }

    private static String normalizeFIO(String fio) {
        if (fio == null) return "";
        return fio.trim()
                .replaceAll("\\s+", " ")
                .replaceAll("[^а-яА-ЯёЁ\\s]", "")
                .toLowerCase();
    }

    private static String normalizeSubject(String subject) {
        if (subject == null || subject.trim().isEmpty()) return "";

        String normalized = subject.trim().toLowerCase();

        // Для математики важно сохранить базовая/профильная
        if (normalized.contains("математика")) {
            if (normalized.contains("баз") || normalized.contains("базовая")) {
                return "Математика базовая";
            } else if (normalized.contains("проф") || normalized.contains("профильная")) {
                return "Математика профильная";
            }
            // Если не указано, оставляем как есть
            return subject.trim();
        }

        // Для информатики - приводим к формату из файла учителей
        if (normalized.contains("информатика") || normalized.contains("икт") || normalized.contains("кегэ")) {
            return "Информатика и ИКТ (КЕГЭ)";
        }

        if (normalized.contains("английский")) {
            if (normalized.contains("устн") || normalized.contains("устный")) {
                return "Английский язык (устный)";
            }
            return "Английский язык";
        }

        // Базовые предметы
        if (normalized.contains("история")) return "История";
        if (normalized.contains("русский")) return "Русский язык";
        if (normalized.contains("физика")) return "Физика";
        if (normalized.contains("химия")) return "Химия";
        if (normalized.contains("биология")) return "Биология";
        if (normalized.contains("обществознание") || normalized.contains("общество")) return "Обществознание";
        if (normalized.contains("литература")) return "Литература";
        if (normalized.contains("география")) return "География";

        return subject.trim();
    }


    private static String normalizeClassName(String className) {
        if (className == null) return "";

        // Удаляем все нецифровые символы, кроме букв
        String normalized = className.trim()
                .replaceAll("\\s+", "")
                .replaceAll("[-\\s]", "")  // Удаляем дефисы и пробелы
                .toUpperCase();

        // Добавляем дефис если его нет и есть буква
        if (!normalized.contains("-") && normalized.matches(".*[А-ЯA-Z].*")) {
            // Ищем позицию первой буквы
            int letterPos = 0;
            for (int i = 0; i < normalized.length(); i++) {
                if (Character.isLetter(normalized.charAt(i))) {
                    letterPos = i;
                    break;
                }
            }
            if (letterPos > 0) {
                normalized = normalized.substring(0, letterPos) + "-" + normalized.substring(letterPos);
            }
        }

        return normalized;
    }

    private static String extractSubject(String line) {
        if (line == null || line.isEmpty()) return "Неизвестный";

        line = line.toLowerCase();

        // Для математики важно сохранять базовая/профильная
        if (line.contains("математика")) {
            if (line.contains("баз") || line.contains("базовая")) {
                return "Математика базовая";
            } else if (line.contains("проф") || line.contains("профильная")) {
                return "Математика профильная";
            }
            return "Математика";
        }

        // Остальные предметы как раньше
        if (line.contains("история")) return "История";
        if (line.contains("русский") || line.contains("русск")) return "Русский язык";
        if (line.contains("физика")) return "Физика";
        if (line.contains("химия")) return "Химия";
        if (line.contains("биология")) return "Биология";
        if (line.contains("обществознание") || line.contains("общество")) return "Обществознание";
        if (line.contains("литература")) return "Литература";
        if (line.contains("география")) return "География";
        if (line.contains("английский")) return "Английский язык";
        if (line.contains("информатика") || line.contains("информатике") || line.contains("информатик") || line.contains("икт") || line.contains("кегэ"))
            return "Информатика и ИКТ (КЕГЭ)";

        return "Неизвестный";
    }

    private static String extractSubjectFromFileName(String fileName) {
        String lowerName = fileName.toLowerCase();

        // Для математики
        if (lowerName.contains("математика")) {
            if (lowerName.contains("баз") || lowerName.contains("базовая")) {
                return "Математика базовая";
            } else if (lowerName.contains("проф") || lowerName.contains("профильная")) {
                return "Математика профильная";
            }
            return "Математика";
        }

        // Остальные предметы
        if (lowerName.contains("история")) return "История";
        if (lowerName.contains("русск")) return "Русский язык";
        if (lowerName.contains("физика")) return "Физика";
        if (lowerName.contains("химия")) return "Химия";
        if (lowerName.contains("биолог")) return "Биология";
        if (lowerName.contains("общест")) return "Обществознание";
        if (lowerName.contains("литература")) return "Литература";
        if (lowerName.contains("географ")) return "География";
        if (lowerName.contains("английск")) return "Английский язык";
        if (lowerName.contains("информат") || lowerName.contains("икт") || lowerName.contains("кегэ"))
            return "Информатика и ИКТ (КЕГЭ)";

        return "Неизвестный предмет";
    }

    private static List<EGKRData> processAllEGKRFiles(SchoolConfig config) throws Exception {
        List<EGKRData> allData = new ArrayList<>();
        File folder = new File(config.egkrFolder);

        if (!folder.exists() || !folder.isDirectory()) {
            throw new FileNotFoundException("Папка не найдена: " + config.egkrFolder);
        }

        File[] files = folder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx") &&
                        !name.toLowerCase().contains("админ") &&
                        !name.toLowerCase().contains("сводка") &&
                        !name.toLowerCase().contains("отчёт"));

        if (files == null || files.length == 0) {
            throw new FileNotFoundException("В папке нет файлов ЕГКР: " + config.egkrFolder);
        }

        System.out.println("\n=== ОБРАБОТКА ФАЙЛОВ ЕГКР ===");
        System.out.println("Найдено " + files.length + " файлов");

        for (File file : files) {
            try {
                System.out.println("\nОбработка: " + file.getName());
                List<EGKRData> fileData = processEGKRFile(file);
                allData.addAll(fileData);
                System.out.println("  Извлечено " + fileData.size() + " записей");
            } catch (Exception e) {
                System.err.println("  Ошибка: " + e.getMessage());
            }
        }

        System.out.println("\nВсего извлечено " + allData.size() + " записей из всех файлов ЕГКР");
        return allData;
    }

    private static List<EGKRData> processEGKRFile(File file) throws Exception {
        List<EGKRData> results = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            // Определяем предмет, дату из ячейки A5 (или подобной)
            String subject = "Неизвестный предмет";
            String date = "2025.12.11";

            // Ищем информацию о предмете и дате в первых строках
            for (int i = 0; i <= 10; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        String cellValue = getCellValue(cell, evaluator);
                        if (!cellValue.isEmpty()) {
                            // Извлекаем предмет
                            String foundSubject = extractSubject(cellValue);
                            if (!foundSubject.equals("Неизвестный")) {
                                subject = foundSubject;
                            }

                            // Извлекаем дату
                            String foundDate = extractDateFromLine(cellValue);
                            if (!foundDate.isEmpty()) {
                                date = foundDate;
                            }
                        }
                    }
                }
            }

            if (subject.equals("Неизвестный предмет")) {
                // Пробуем извлечь из имени файла
                subject = extractSubjectFromFileName(file.getName());
            }

            // Находим строку с заголовками таблицы
            int dataStartRow = findDataStartRow(sheet);
            if (dataStartRow == -1) {
                return results;
            }

            // Определяем индексы колонок
            Map<String, Integer> columnIndexes = findColumnIndexes(sheet, dataStartRow, evaluator);

            // Обрабатываем строки с данными
            for (int i = dataStartRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Пропускаем строки с итогами
                if (isTotalRow(row, evaluator)) {
                    continue;
                }

                // Извлекаем данные
                EGKRData data = extractDataFromRow(row, columnIndexes, subject, date, evaluator);
                if (data != null) {
                    results.add(data);
                }
            }
        }

        return results;
    }

    private static int findDataStartRow(Sheet sheet) {
        // Ищем строку с заголовком "№ п/п" или подобным
        for (int i = 0; i <= 20; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null) {
                    String cellValue = getCellValue(cell).trim();
                    // Проверяем различные варианты написания заголовка
                    if (cellValue.equals("№ п/п") ||
                            cellValue.equals("№") ||
                            cellValue.startsWith("№ п") ||
                            cellValue.equalsIgnoreCase("№ п/п") ||
                            cellValue.contains("№") && (cellValue.contains("п/п") || cellValue.contains("п/п"))) {
                        return i;
                    }
                }
            }
        }

        // Если не нашли стандартный заголовок, ищем по другим признакам
        for (int i = 0; i <= 20; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                // Проверяем, содержит ли строка несколько заголовков
                int headerCount = 0;
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        String cellValue = getCellValue(cell).trim().toLowerCase();
                        if (cellValue.contains("фамилия") ||
                                cellValue.contains("имя") ||
                                cellValue.contains("класс") ||
                                cellValue.contains("код оо") ||
                                cellValue.contains("фио")) {
                            headerCount++;
                        }
                    }
                }
                if (headerCount >= 2) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static Map<String, Integer> findColumnIndexes(Sheet sheet, int headerRow, FormulaEvaluator evaluator) {
        Map<String, Integer> indexes = new HashMap<>();
        Row header = sheet.getRow(headerRow);

        if (header == null) return indexes;

        for (int i = 0; i < header.getLastCellNum(); i++) {
            Cell cell = header.getCell(i);
            if (cell != null) {
                String headerText = getCellValue(cell, evaluator).trim();
                if (!headerText.isEmpty()) {
                    indexes.put(headerText, i);

                    // Также добавляем нормализованные версии заголовков
                    String normalized = headerText.toLowerCase();
                    if (!indexes.containsKey(normalized)) {
                        indexes.put(normalized, i);
                    }
                }
            }
        }

        return indexes;
    }

    private static boolean isTotalRow(Row row, FormulaEvaluator evaluator) {
        if (row == null) return true;

        Cell firstCell = row.getCell(0);
        if (firstCell != null) {
            String cellValue = getCellValue(firstCell, evaluator).toLowerCase();
            return cellValue.contains("всего") ||
                    cellValue.contains("итог") ||
                    cellValue.contains("итого") ||
                    cellValue.contains("участник");
        }

        return false;
    }

    private static EGKRData extractDataFromRow(Row row, Map<String, Integer> columnIndexes,
                                               String subject, String date, FormulaEvaluator evaluator) {
        try {
            // Получаем класс (ищем по разным возможным названиям)
            String className = "";
            Integer classIndex = findColumnIndex(columnIndexes,
                    new String[]{"Класс", "класс"});
            if (classIndex != null) {
                className = getCellValue(row.getCell(classIndex), evaluator);
            }

            // Получаем ФИО
            String lastName = "";
            String firstName = "";
            String middleName = "";

            // Вариант 1: отдельные колонки
            Integer lastNameIndex = findColumnIndex(columnIndexes,
                    new String[]{"Фамилия", "фамилия"});
            Integer firstNameIndex = findColumnIndex(columnIndexes,
                    new String[]{"Имя", "имя"});
            Integer middleNameIndex = findColumnIndex(columnIndexes,
                    new String[]{"Отчество", "отчество"});

            if (lastNameIndex != null && firstNameIndex != null) {
                lastName = getCellValue(row.getCell(lastNameIndex), evaluator);
                firstName = getCellValue(row.getCell(firstNameIndex), evaluator);
                if (middleNameIndex != null) {
                    middleName = getCellValue(row.getCell(middleNameIndex), evaluator);
                }
            } else {
                // Вариант 2: ФИО в одной колонке
                Integer fioIndex = findColumnIndex(columnIndexes,
                        new String[]{"ФИО", "фио", "Фамилия Имя Отчество"});
                if (fioIndex != null) {
                    String fullName = getCellValue(row.getCell(fioIndex), evaluator);
                    String[] nameParts = fullName.split("\\s+");
                    if (nameParts.length >= 1) lastName = nameParts[0];
                    if (nameParts.length >= 2) firstName = nameParts[1];
                    if (nameParts.length >= 3) middleName = nameParts[2];
                }
            }

            if (lastName.isEmpty() || firstName.isEmpty()) {
                return null;
            }

            // Получаем результаты заданий
            String shortAnswerTasks = "";
            String longAnswerTasks = "";

            Integer shortAnswerIndex = findColumnIndex(columnIndexes,
                    new String[]{"Задания с кратким ответом"});
            if (shortAnswerIndex != null) {
                shortAnswerTasks = getCellValue(row.getCell(shortAnswerIndex), evaluator);
            }

            Integer longAnswerIndex = findColumnIndex(columnIndexes,
                    new String[]{"Задания с развернутым ответом"});
            if (longAnswerIndex != null) {
                longAnswerTasks = getCellValue(row.getCell(longAnswerIndex), evaluator);
            }

            // Получаем первичный балл
            double primaryScore = 0;
            Integer scoreIndex = findColumnIndex(columnIndexes,
                    new String[]{"Первичный балл", "балл", "Балл"});

            if (scoreIndex != null) {
                String scoreStr = getCellValue(row.getCell(scoreIndex), evaluator);
                try {
                    if (!scoreStr.isEmpty()) {
                        primaryScore = Double.parseDouble(scoreStr.replace(",", "."));
                    }
                } catch (NumberFormatException e) {
                    Cell scoreCell = row.getCell(scoreIndex);
                    if (scoreCell != null && scoreCell.getCellType() == CellType.NUMERIC) {
                        primaryScore = scoreCell.getNumericCellValue();
                    }
                }
            }

            // Создаем объект данных
            return new EGKRData(
                    subject,
                    date,
                    className,
                    lastName,
                    firstName,
                    middleName,
                    shortAnswerTasks,
                    longAnswerTasks,
                    primaryScore,
                    0
            );

        } catch (Exception e) {
            return null;
        }
    }

    private static boolean isValidData(String fio, String teacher) {
        if (fio == null || fio.trim().isEmpty()) return false;
        if (teacher == null || teacher.trim().isEmpty()) return false;

        String teacherLower = teacher.toLowerCase();
        return !teacherLower.equals("нет") &&
                !teacherLower.equals("не указан") &&
                !teacherLower.equals("неизвестно") &&
                !teacherLower.equals("null") &&
                !teacherLower.equals("-") &&
                !teacherLower.equals("—") &&
                !teacherLower.equals("отсутствует");
    }

    /**
     * Улучшенный метод получения значения ячейки с поддержкой формул
     * Сначала пытается получить вычисленное значение, затем сырое
     */
    private static String getCellValue(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return "";

        try {
            // Если это формула, пытаемся вычислить ее значение
            if (cell.getCellType() == CellType.FORMULA) {
                try {
                    // Пытаемся вычислить формулу
                    CellValue cellValue = evaluator.evaluate(cell);
                    if (cellValue != null) {
                        switch (cellValue.getCellType()) {
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    return new SimpleDateFormat("yyyy.MM.dd").format(cell.getDateCellValue());
                                } else {
                                    double value = cellValue.getNumberValue();
                                    if (value == Math.floor(value)) {
                                        return String.valueOf((int) value);
                                    } else {
                                        return String.valueOf(value);
                                    }
                                }
                            case STRING:
                                return cellValue.getStringValue().trim();
                            case BOOLEAN:
                                return String.valueOf(cellValue.getBooleanValue());
                            case ERROR:
                                // Если ошибка при вычислении, возвращаем формулу
                                return cell.getCellFormula();
                            default:
                                return "";
                        }
                    }
                } catch (Exception e) {
                    // Если не удалось вычислить, возвращаем формулу как строку
                    return cell.getCellFormula();
                }
            }

            // Для обычных ячеек используем стандартный метод
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return new SimpleDateFormat("yyyy.MM.dd").format(cell.getDateCellValue());
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
                    // Дублируем обработку формул для надежности
                    return cell.getCellFormula();
                default:
                    return "";
            }
        } catch (Exception e) {
            // В случае любой ошибки возвращаем пустую строку
            return "";
        }
    }

    /**
     * Перегруженный метод для случаев, когда evaluator не нужен
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";

        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return new SimpleDateFormat("yyyy.MM.dd").format(cell.getDateCellValue());
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
                    // Без evaluator просто возвращаем формулу
                    return cell.getCellFormula();
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    @Data
    static class EGKRData {
        private String subject;
        private String date;
        private String className;
        private String lastName;
        private String firstName;
        private String middleName;
        private String shortAnswerTasks;
        private String longAnswerTasks;
        private double primaryScore;
        private double percent;
        private String teacherName;

        public EGKRData(String subject, String date, String className,
                        String lastName, String firstName, String middleName,
                        String shortAnswerTasks, String longAnswerTasks,
                        double primaryScore, double percent) {
            this.subject = subject;
            this.date = date;
            this.className = className;
            this.lastName = lastName;
            this.firstName = firstName;
            this.middleName = middleName;
            this.shortAnswerTasks = shortAnswerTasks;
            this.longAnswerTasks = longAnswerTasks;
            this.primaryScore = primaryScore;
            this.percent = percent;
        }

        public String getFullName() {
            return lastName + " " + firstName + (middleName != null && !middleName.isEmpty() ? " " + middleName : "");
        }
    }

    private static Integer findColumnIndex(Map<String, Integer> columnIndexes, String[] possibleNames) {
        for (String name : possibleNames) {
            // Прямой поиск
            if (columnIndexes.containsKey(name)) {
                return columnIndexes.get(name);
            }

            // Поиск по частичному совпадению
            for (String key : columnIndexes.keySet()) {
                String normalizedKey = key.toLowerCase().replaceAll("\\s+", "");
                String normalizedName = name.toLowerCase().replaceAll("\\s+", "");

                if (normalizedKey.contains(normalizedName) || normalizedName.contains(normalizedKey)) {
                    return columnIndexes.get(key);
                }
            }
        }

        return null;
    }

    private static int findHeaderRow(Sheet sheet) {
        // Проверяем первые 20 строк
        for (int i = 0; i <= 20; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int foundColumns = 0;

                for (Cell cell : row) {
                    if (cell != null) {
                        String cellValue = cell.toString().trim().toLowerCase();
                        if (!cellValue.isEmpty()) {
                            // Проверяем на наличие ключевых слов заголовков
                            if (cellValue.contains("фио") ||
                                    cellValue.contains("класс") ||
                                    cellValue.contains("учитель") ||
                                    cellValue.contains("предмет") ||
                                    cellValue.contains("ученик")) {
                                foundColumns++;
                            }
                        }
                    }
                }

                // Если найдено хотя бы 2 ключевых слова, считаем это заголовком
                if (foundColumns >= 2) {
                    return i;
                }
            }
        }

        // Если не нашли, возвращаем первую строку
        return 0;
    }


}

// ВАЖНО: Вам нужно скопировать ВСЕ остальные методы из оригинального класса EGKRReportGenerator
// и вставить их в этот новый класс EGKRReportGeneratorV2