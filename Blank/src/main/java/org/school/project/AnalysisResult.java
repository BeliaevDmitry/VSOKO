package org.school.project;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;

public class AnalysisResult {

    // ===== НАСТРОЙКИ ПУТЕЙ =====
    private static final String INPUT_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\На разбор";
    private static final String REPORTS_BASE_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО\\Работы\\{предмет}\\Отчёты";
    private static final String FINAL_REPORT_FOLDER = "C:\\Users\\dimah\\Yandex.Disk\\ГБОУ 7\\ВСОКО";

    // ===== ОСНОВНОЙ МЕТОД ДЛЯ ЗАПУСКА =====
    public static void main(String[] args) {
        try {
            System.out.println("=== Запуск анализа отчётов ===");
            System.out.println("Папка с исходными файлами: " + INPUT_FOLDER);

            // Запускаем обработку всех файлов
            List<StudentResult> allResults = processAllReports();

            // Выводим сводную информацию
            System.out.println("\n=== РЕЗУЛЬТАТЫ АНАЛИЗА ===");
            System.out.println("Всего обработано результатов: " + allResults.size());

        } catch (Exception e) {
            System.err.println("Ошибка при выполнении анализа: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // ===== МЕТОД 1 =====
    /**
     * Основной метод обработки всех отчётов
     *
     * Этот метод выполняет следующие действия:
     * 1. Находит все Excel файлы в указанной папке
     * 2. Для каждого файла вызывает парсер для извлечения данных
     * 3. Перемещает обработанные файлы в папки по предметам
     * 4. Возвращает объединённый список всех результатов
     *
     * @return Список объектов StudentResult со всеми данными из всех файлов
     * @throws IOException Если возникают ошибки чтения/записи файлов
     */
    public static List<StudentResult> processAllReports() throws IOException {
        System.out.println("\n=== НАЧАЛО ОБРАБОТКИ ВСЕХ ОТЧЁТОВ ===");

        List<StudentResult> allResults = new ArrayList<>();

        // Шаг 1: Получаем список всех Excel файлов в папке
        File[] reportFiles = getExcelFilesFromFolder();

        if (reportFiles == null || reportFiles.length == 0) {
            System.out.println("В папке " + INPUT_FOLDER + " не найдено файлов для обработки");
            return allResults;
        }

        System.out.println("Найдено файлов для обработки: " + reportFiles.length);

        // Шаг 2: Обрабатываем каждый файл по очереди
        for (int i = 0; i < reportFiles.length; i++) {
            File reportFile = reportFiles[i];

            try {
                System.out.printf("\n[%d/%d] Обработка файла: %s%n",
                        i + 1, reportFiles.length, reportFile.getName());

                // Шаг 3: Парсим файл и извлекаем данные учеников
                List<StudentResult> fileResults = parseReportFile(reportFile);

                if (fileResults.isEmpty()) {
                    System.out.println("  ⚠ В файле не найдено данных учеников");
                    continue;
                }

                // Шаг 4: Добавляем результаты в общий список
                allResults.addAll(fileResults);
                System.out.printf("  ✓ Извлечено результатов: %d%n", fileResults.size());

                // Шаг 5: Определяем предмет для организации файлов
                String subject = fileResults.get(0).getSubject();

                // Шаг 6: Перемещаем файл в соответствующую папку предмета
                moveReportToSubjectFolder(reportFile, subject);
                System.out.printf("  ✓ Файл перемещён в папку предмета: %s%n", subject);

            } catch (Exception e) {
                System.err.printf("  ✗ Ошибка при обработке файла %s: %s%n",
                        reportFile.getName(), e.getMessage());
            }
        }

        System.out.println("\n=== ОБРАБОТКА ЗАВЕРШЕНА ===");
        System.out.println("Всего обработано учеников: " + allResults.size());

        return allResults;
    }

    // ===== МЕТОД 2 =====
    /**
     * Парсинг одного файла отчёта и извлечение данных учеников
     *
     * Этот метод анализирует Excel файл, созданный DataCollectionSheetGenerator,
     * и извлекает следующие данные для каждого ученика:
     * - Класс и предмет (из листа "Информация")
     * - ФИО ученика
     * - Присутствие на тестировании
     * - Номер варианта работы
     * - Баллы за каждое задание (количество заданий определяется автоматически)
     * - Итоговый балл (сумма всех баллов)
     *
     * Особенности:
     * - Обрабатывает максимум 34 ученика на класс
     * - Автоматически определяет количество заданий
     * - Пропускает отсутствующих учеников (помеченных как "Не был")
     * - Поддерживает разное количество заданий в разных работах
     *
     * @param reportFile Файл отчёта в формате .xlsx
     * @return Список объектов StudentResult с данными учеников из файла
     * @throws IOException Если файл не найден или имеет неверный формат
     */
    public static List<StudentResult> parseReportFile(File reportFile) throws IOException {
        List<StudentResult> results = new ArrayList<>();

        // Открываем Excel файл для чтения
        try (FileInputStream file = new FileInputStream(reportFile);
             Workbook workbook = new XSSFWorkbook(file)) {

            // === ШАГ 1: Получаем общую информацию (предмет и класс) ===
            Sheet infoSheet = workbook.getSheet("Информация");
            String subject = "Неизвестный предмет";
            String className = "Неизвестный класс";

            if (infoSheet != null) {
                // Предмет находится в строке 3, колонка B (0-based: row=2, col=1)
                subject = getCellValueAsString(infoSheet, 2, 1, "Неизвестный предмет");

                // Класс находится в строке 4, колонка B (0-based: row=3, col=1)
                className = getCellValueAsString(infoSheet, 3, 1, "Неизвестный класс");
            }

            // === ШАГ 2: Получаем данные учеников ===
            Sheet dataSheet = workbook.getSheet("Сбор информации");
            if (dataSheet == null) {
                throw new IOException("В файле отсутствует лист 'Сбор информации'");
            }

            // Определяем количество заданий в работе
            int taskCount = determineTaskCount(dataSheet);
            System.out.printf("  Обнаружено заданий в работе: %d%n", taskCount);

            // === ШАГ 3: Обрабатываем данные учеников (начинаются с 4 строки) ===
            // В Excel строки: 1-заголовок, 2-номера заданий, 3-макс. баллы, 4-данные учеников
            int firstStudentRow = 3; // 4-я строка в Excel (0-based индекс)
            int maxStudentsToProcess = 34; // Максимальное количество учеников в классе

            for (int rowIndex = firstStudentRow;
                 rowIndex < firstStudentRow + maxStudentsToProcess;
                 rowIndex++) {

                Row row = dataSheet.getRow(rowIndex);
                if (row == null) {
                    // Достигли конца списка учеников
                    break;
                }

                // === ШАГ 4: Извлекаем ФИО ученика ===
                String fio = getCellValueAsString(row.getCell(1)); // Колонка B (ФИО)
                if (fio == null || fio.trim().isEmpty()) {
                    // Пустая строка - возможно, конец списка или пустая запись
                    continue;
                }

                // === ШАГ 5: Проверяем присутствие ученика ===
                String presence = getCellValueAsString(row.getCell(2)); // Колонка C (Присутствие)
                if ("Не был".equalsIgnoreCase(presence)) {
                    // Пропускаем отсутствующих учеников
                    continue;
                }

                // === ШАГ 6: Получаем номер варианта ===
                String variant = getCellValueAsString(row.getCell(3)); // Колонка D (Вариант)

                // === ШАГ 7: Создаём объект результата ===
                StudentResult result = new StudentResult();
                result.setSubject(subject);
                result.setClassName(className);
                result.setFio(fio.trim());
                result.setPresence(presence);
                result.setVariant(variant);

                // === ШАГ 8: Собираем баллы за задания ===
                Map<Integer, Integer> taskScores = new HashMap<>();
                int totalScore = 0;

                // Баллы за задания начинаются с колонки E (индекс 4)
                for (int taskNum = 1; taskNum <= taskCount; taskNum++) {
                    int columnIndex = 3 + taskNum; // E=4, F=5, и т.д.
                    Cell scoreCell = row.getCell(columnIndex);

                    // Получаем балл за задание
                    Integer score = getCellValueAsInteger(scoreCell);
                    if (score == null) {
                        score = 0; // Если ячейка пустая, считаем 0 баллов
                    }

                    taskScores.put(taskNum, score);
                    totalScore += score;
                }

                // === ШАГ 9: Сохраняем результаты ===
                result.setTaskScores(taskScores);
                result.setTotalScore(totalScore);
                results.add(result);
            }

        } catch (Exception e) {
            throw new IOException("Ошибка при чтении файла " + reportFile.getName(), e);
        }

        return results;
    }

    // ===== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ =====

    /**
     * Получает список всех Excel файлов из папки INPUT_FOLDER
     *
     * @return Массив файлов с расширением .xlsx или пустой массив, если файлов нет
     */
    private static File[] getExcelFilesFromFolder() {
        File inputFolder = new File(INPUT_FOLDER);

        if (!inputFolder.exists()) {
            System.err.println("Ошибка: Папка не существует - " + INPUT_FOLDER);
            return new File[0];
        }

        if (!inputFolder.isDirectory()) {
            System.err.println("Ошибка: Указанный путь не является папкой - " + INPUT_FOLDER);
            return new File[0];
        }

        // Фильтруем только файлы с расширением .xlsx
        return inputFolder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx")
        );
    }

    /**
     * Перемещает обработанный файл в папку соответствующего предмета
     *
     * Создаёт структуру папок: {REPORTS_BASE_FOLDER}/{предмет}/Отчёты/
     * и перемещает туда файл
     *
     * @param reportFile Файл для перемещения
     * @param subject Название предмета (используется для создания папки)
     * @throws IOException Если не удалось создать папку или переместить файл
     */
    private static void moveReportToSubjectFolder(File reportFile, String subject) throws IOException {
        // Если предмет не определён, пытаемся извлечь из имени файла
        if (subject == null || subject.isEmpty() || "Неизвестный предмет".equals(subject)) {
            subject = extractSubjectFromFileName(reportFile.getName());
        }

        // Очищаем название предмета от недопустимых символов для имени папки
        String safeSubject = subject.replaceAll("[\\\\/:*?\"<>|]", "_");

        // Формируем путь к целевой папке
        String subjectFolderPath = REPORTS_BASE_FOLDER.replace("{предмет}", safeSubject);
        Path subjectFolder = Paths.get(subjectFolderPath);

        // Создаём целевую папку, если она не существует
        if (!Files.exists(subjectFolder)) {
            Files.createDirectories(subjectFolder);
        }

        // Перемещаем файл
        Path source = reportFile.toPath();
        Path target = subjectFolder.resolve(reportFile.getName());

        try {
            Files.move(source, target, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            System.err.println("Ошибка при перемещении файла " + reportFile.getName() +
                    " в папку " + subjectFolder);
            throw e;
        }
    }

    /**
     * Извлекает название предмета из имени файла
     *
     * Анализирует имя файла в формате "Сбор_данных_10А_История.xlsx"
     * и пытается найти название предмета
     *
     * @param fileName Имя файла
     * @return Название предмета или "Неизвестный предмет", если не удалось определить
     */
    private static String extractSubjectFromFileName(String fileName) {
        // Пример формата: "Сбор_данных_10А_История.xlsx"
        String[] parts = fileName.split("_");

        // Ищем последнюю часть перед расширением, которая не является номером класса
        for (int i = parts.length - 1; i >= 0; i--) {
            String part = parts[i].replace(".xlsx", "").trim();

            // Пропускаем служебные слова и номера классов
            if (!part.isEmpty() &&
                    !part.matches("\\d+[А-Яа-я]?") &&  // Не номер класса (10, 10А и т.д.)
                    !part.equalsIgnoreCase("Сбор") &&
                    !part.equalsIgnoreCase("данных") &&
                    !part.equalsIgnoreCase("класс")) {
                return part;
            }
        }

        return "Неизвестный предмет";
    }

    /**
     * Определяет количество заданий в работе
     *
     * Анализирует вторую строку листа "Сбор информации",
     * где в ячейках находятся номера заданий (1, 2, 3, ...)
     *
     * @param sheet Лист "Сбор информации"
     * @return Количество заданий в работе
     */
    private static int determineTaskCount(Sheet sheet) {
        // Номера заданий находятся на второй строке (индекс 1)
        Row taskNumberRow = sheet.getRow(1);
        if (taskNumberRow == null) {
            return 0;
        }

        int taskCount = 0;

        // Начинаем с колонки E (индекс 4) и идём до тех пор, пока находим числа
        for (int col = 4; col < 100; col++) { // Максимум 100 заданий
            Cell cell = taskNumberRow.getCell(col);
            if (cell == null) {
                break; // Ячейка пустая - конец списка заданий
            }

            String value = getCellValueAsString(cell);
            if (value != null && value.matches("\\d+")) {
                taskCount++; // Нашли номер задания
            } else {
                break; // Нашли не-число - конец списка заданий
            }
        }

        return taskCount;
    }

    /**
     * Получает строковое значение из ячейки по координатам
     *
     * @param sheet Лист Excel
     * @param rowIndex Индекс строки (0-based)
     * @param colIndex Индекс колонки (0-based)
     * @param defaultValue Значение по умолчанию, если ячейка не найдена
     * @return Значение ячейки или defaultValue
     */
    private static String getCellValueAsString(Sheet sheet, int rowIndex, int colIndex, String defaultValue) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return defaultValue;
        }

        Cell cell = row.getCell(colIndex);
        return getCellValueAsString(cell, defaultValue);
    }

    /**
     * Получает строковое значение из ячейки
     *
     * @param cell Ячейка Excel
     * @return Значение ячейки как строка или null
     */
    private static String getCellValueAsString(Cell cell) {
        return getCellValueAsString(cell, null);
    }

    /**
     * Получает строковое значение из ячейки с указанием значения по умолчанию
     *
     * @param cell Ячейка Excel
     * @param defaultValue Значение по умолчанию, если ячейка пустая
     * @return Значение ячейки как строка
     */
    private static String getCellValueAsString(Cell cell, String defaultValue) {
        if (cell == null) {
            return defaultValue;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    // Если целое число - убираем дробную часть
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
                return defaultValue;
        }
    }

    /**
     * Получает целочисленное значение из ячейки
     *
     * @param cell Ячейка Excel
     * @return Значение ячейки как Integer или null, если не удалось преобразовать
     */
    private static Integer getCellValueAsInteger(Cell cell) {
        String stringValue = getCellValueAsString(cell);
        if (stringValue == null || stringValue.trim().isEmpty()) {
            return null;
        }

        try {
            // Убираем возможную дробную часть (.0)
            stringValue = stringValue.replaceAll("\\.0+$", "");
            return Integer.parseInt(stringValue);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    // ===== КЛАСС ДЛЯ ХРАНЕНИЯ РЕЗУЛЬТАТОВ УЧЕНИКА =====

    /**
     * Класс для хранения всех данных об ученике из отчёта
     *
     * Содержит полную информацию о результатах тестирования одного ученика:
     * - Основные данные (предмет, класс, ФИО)
     * - Информация о тестировании (присутствие, вариант)
     * - Результаты (баллы по заданиям, итоговый балл)
     */
    public static class StudentResult {
        private String subject;                    // Предмет тестирования
        private String className;                  // Класс ученика
        private String fio;                        // ФИО ученика
        private String presence;                   // Присутствие ("Был"/"Не был")
        private String variant;                    // Номер варианта работы
        private Map<Integer, Integer> taskScores;  // Баллы по заданиям (номер → балл)
        private int totalScore;                    // Итоговый балл

        public StudentResult() {
            this.taskScores = new HashMap<>();
        }

        // Геттеры и сеттеры
        public String getSubject() { return subject; }
        public void setSubject(String subject) { this.subject = subject; }

        public String getClassName() { return className; }
        public void setClassName(String className) { this.className = className; }

        public String getFio() { return fio; }
        public void setFio(String fio) { this.fio = fio; }

        public String getPresence() { return presence; }
        public void setPresence(String presence) { this.presence = presence; }

        public String getVariant() { return variant; }
        public void setVariant(String variant) { this.variant = variant; }

        public Map<Integer, Integer> getTaskScores() { return taskScores; }
        public void setTaskScores(Map<Integer, Integer> taskScores) {
            this.taskScores = taskScores;
        }

        public int getTotalScore() { return totalScore; }
        public void setTotalScore(int totalScore) { this.totalScore = totalScore; }

        @Override
        public String toString() {
            return String.format("StudentResult{subject='%s', className='%s', fio='%s', " +
                            "totalScore=%d, tasks=%d}",
                    subject, className, fio, totalScore, taskScores.size());
        }
    }
}