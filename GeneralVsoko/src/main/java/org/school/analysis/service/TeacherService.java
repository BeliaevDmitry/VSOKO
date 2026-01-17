package org.school.analysis.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.school.analysis.config.AppConfig;
import org.school.analysis.model.Teacher;
import org.school.analysis.repository.TeacherRepository;
import org.school.analysis.util.TeacherNameNormalizer;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import jakarta.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDateTime;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

@Service
@Slf4j
@RequiredArgsConstructor
public class TeacherService {

    private final TeacherRepository teacherRepository;

    private final Map<String, Teacher> teachersCache = new ConcurrentHashMap<>();
    private final Map<String, String> normalizedToFullName = new ConcurrentHashMap<>();
    private final Map<String, List<String>> lastNameVariants = new ConcurrentHashMap<>();

    private String lastFileHash = "";
    private LocalDateTime lastImportTime;


       /**
     * Расчет хэша файла для отслеживания изменений
     */
    private String calculateFileHash(File file) {
        try {
            return String.valueOf(file.lastModified() + file.length());
        } catch (Exception e) {
            log.warn("Не удалось рассчитать хэш файла: {}", e.getMessage());
            return "";
        }
    }

    /**
     * Импорт учителей из Excel файла
     */
    @Transactional
    public void importTeachersFromExcel(String school) {
        String filePath = AppConfig.INPUT_TEACHER_NAME.replace("{школа}", school);
        File teacherFile = new File(filePath);

        if (!teacherFile.exists()) {
            log.warn("Файл учителей не найден: {}. Пропускаем импорт.", filePath);
            return;
        }

        log.info("Начинаем импорт учителей из файла: {}", filePath);

        try (FileInputStream file = new FileInputStream(teacherFile);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);
            int importedCount = 0;
            int skippedCount = 0;
            int errorCount = 0;

            for (Row row : sheet) {
                // Пропускаем заголовок (первая строка)
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell nameCell = row.getCell(1); // Столбец B
                if (nameCell == null || isCellEmpty(nameCell)) {
                    continue;
                }

                String fullName = getCellValue(nameCell).trim();
                if (fullName.isEmpty()) {
                    continue;
                }

                try {
                    Teacher teacher = createOrUpdateTeacher(fullName, filePath);
                    if (teacher != null) {
                        importedCount++;
                    } else {
                        skippedCount++;
                    }
                } catch (Exception e) {
                    log.error("Ошибка при импорте учителя '{}': {}", fullName, e.getMessage());
                    errorCount++;
                }
            }

            // Обновляем кэш после импорта
            loadTeachersFromDatabase();

            // Сохраняем информацию об импорте
            lastFileHash = calculateFileHash(teacherFile);
            lastImportTime = LocalDateTime.now();

            log.info("Импорт учителей завершен. " +
                            "Импортировано: {}, Пропущено: {}, Ошибок: {}",
                    importedCount, skippedCount, errorCount);

        } catch (Exception e) {
            log.error("Ошибка при импорте учителей из файла: {}", e.getMessage(), e);
        }
    }

    /**
     * Создает или обновляет учителя в БД
     */
    @Transactional
    public Teacher createOrUpdateTeacher(String fullName, String sourceFile) {
        String normalized = TeacherNameNormalizer.normalize(fullName);

        // Проверяем, существует ли уже такой учитель
        Optional<Teacher> existingTeacher = teacherRepository.findByNormalizedFullName(normalized);

        if (existingTeacher.isPresent()) {
            // Обновляем существующего учителя
            Teacher teacher = existingTeacher.get();
            if (!teacher.getIsActive()) {
                teacher.setIsActive(true);
                teacher.setSourceFile(sourceFile);
                teacher.setUpdatedAt(LocalDateTime.now());
                teacherRepository.save(teacher);
                log.debug("Активирован существующий учитель: {}", fullName);
            }
            return teacher;
        }

        // Создаем нового учителя
        String[] parts = normalized.split(" ");
        if (parts.length < 2) {
            log.warn("Некорректное ФИО для импорта: {}", fullName);
            return null;
        }

        Teacher teacher = Teacher.builder()
                .fullName(fullName)
                .normalizedFullName(normalized)
                .lastName(parts[0])
                .firstName(parts.length > 1 ? parts[1] : "")
                .middleName(parts.length > 2 ? parts[2] : "")
                .shortName(TeacherNameNormalizer.getShortName(fullName))
                .isActive(true)
                .sourceFile(sourceFile)
                .build();

        Teacher saved = teacherRepository.save(teacher);
        log.debug("Создан новый учитель: {}", fullName);
        return saved;
    }

    /**
     * Загрузка учителей из БД в кэш
     */
    private void loadTeachersFromDatabase() {
        log.info("Загрузка учителей из базы данных...");

        teachersCache.clear();
        normalizedToFullName.clear();
        lastNameVariants.clear();

        List<Teacher> activeTeachers = teacherRepository.findByIsActiveTrue();

        for (Teacher teacher : activeTeachers) {
            teachersCache.put(teacher.getFullName(), teacher);
            normalizedToFullName.put(teacher.getNormalizedFullName(), teacher.getFullName());

            // Добавляем варианты для поиска по фамилии
            if (teacher.getLastName() != null && !teacher.getLastName().isEmpty()) {
                lastNameVariants
                        .computeIfAbsent(teacher.getLastName().toLowerCase(), k -> new ArrayList<>())
                        .add(teacher.getFullName());
            }
        }

        log.info("Загружено {} активных учителей из БД", activeTeachers.size());
    }

    /**
     * Обновление кэша учителей
     */
    @Scheduled(fixedDelay = 300000) // Каждые 5 минут
    public void refreshCache() {
        log.debug("Обновление кэша учителей...");
        loadTeachersFromDatabase();
    }

    /**
     * Проверка валидности учителя
     */
    public boolean isTeacherValid(String teacherNameFromReport) {
        if (teacherNameFromReport == null || teacherNameFromReport.trim().isEmpty()) {
            log.warn("Имя учителя пустое");
            return false;
        }

        String normalized = TeacherNameNormalizer.normalize(teacherNameFromReport);

        // 1. Прямое совпадение
        if (normalizedToFullName.containsKey(normalized)) {
            updateTeacherLastSeen(teacherNameFromReport);
            return true;
        }

        // 2. Поиск по фамилии
        String[] parts = normalized.split(" ");
        if (parts.length > 0) {
            String lastName = parts[0];
            if (lastNameVariants.containsKey(lastName)) {
                updateTeacherLastSeen(teacherNameFromReport);
                return true;
            }
        }

        // 3. Нечеткий поиск в БД
        Optional<Teacher> fuzzyMatch = findFuzzyMatchInDatabase(teacherNameFromReport);
        if (fuzzyMatch.isPresent()) {
            updateTeacherLastSeen(teacherNameFromReport);
            return true;
        }

        log.warn("Учитель не найден в базе: {}", teacherNameFromReport);
        return false;
    }

    /**
     * Нечеткий поиск в БД
     */
    private Optional<Teacher> findFuzzyMatchInDatabase(String teacherName) {
        String normalized = TeacherNameNormalizer.normalize(teacherName);
        String[] parts = normalized.split(" ");

        if (parts.length == 0) {
            return Optional.empty();
        }

        // Ищем по фамилии в БД
        List<Teacher> sameLastNameTeachers = teacherRepository
                .findByLastNameIgnoreCase(parts[0]);

        for (Teacher teacher : sameLastNameTeachers) {
            if (TeacherNameNormalizer.isSimilar(normalized, teacher.getNormalizedFullName())) {
                return Optional.of(teacher);
            }
        }

        return Optional.empty();
    }

    /**
     * Обновление времени последнего использования учителя
     */
    @Transactional
    public void updateTeacherLastSeen(String teacherName) {
        try {
            String normalized = TeacherNameNormalizer.normalize(teacherName);
            Optional<Teacher> teacherOpt = teacherRepository.findByNormalizedFullName(normalized);

            if (teacherOpt.isEmpty()) {
                // Пробуем найти по фамилии
                String[] parts = normalized.split(" ");
                if (parts.length > 0) {
                    List<Teacher> teachers = teacherRepository.findByLastNameIgnoreCase(parts[0]);
                    if (!teachers.isEmpty()) {
                        teacherOpt = Optional.of(teachers.get(0));
                    }
                }
            }

            if (teacherOpt.isPresent()) {
                Teacher teacher = teacherOpt.get();
                teacher.setLastSeenInReport(LocalDateTime.now());
                teacherRepository.save(teacher);
            }
        } catch (Exception e) {
            log.warn("Не удалось обновить lastSeen для учителя {}: {}", teacherName, e.getMessage());
        }
    }

    /**
     * Получение полного ФИО учителя
     */
    public Optional<String> getFullTeacherName(String teacherNameFromReport) {
        if (teacherNameFromReport == null) {
            return Optional.empty();
        }

        String normalized = TeacherNameNormalizer.normalize(teacherNameFromReport);

        // 1. Из кэша
        if (normalizedToFullName.containsKey(normalized)) {
            return Optional.of(normalizedToFullName.get(normalized));
        }

        // 2. Из БД
        Optional<Teacher> teacherOpt = findFuzzyMatchInDatabase(teacherNameFromReport);
        return teacherOpt.map(Teacher::getFullName);
    }

    /**
     * Получение всех активных учителей
     */
    public List<Teacher> getAllActiveTeachers() {
        return new ArrayList<>(teachersCache.values());
    }

    /**
     * Добавление нового учителя вручную
     */
    @Transactional
    public Teacher addTeacherManually(String fullName) {
        Teacher teacher = createOrUpdateTeacher(fullName, "manual_add");
        if (teacher != null) {
            loadTeachersFromDatabase(); // Обновляем кэш
        }
        return teacher;
    }

    /**
     * Деактивация учителя
     */
    @Transactional
    public void deactivateTeacher(Long teacherId) {
        teacherRepository.findById(teacherId).ifPresent(teacher -> {
            teacher.setIsActive(false);
            teacherRepository.save(teacher);
            loadTeachersFromDatabase(); // Обновляем кэш
        });
    }

    /**
     * Получение значения ячейки
     */
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }

    /**
     * Проверка пустая ли ячейка
     */
    private boolean isCellEmpty(Cell cell) {
        if (cell == null) {
            return true;
        }

        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim().isEmpty();
        }

        return false;
    }

    /**
     * Получить статистику по учителям
     */
    public Map<String, Object> getStatistics() {
        Map<String, Object> stats = new HashMap<>();
        stats.put("totalInDb", teacherRepository.count());
        stats.put("activeTeachers", teacherRepository.countActiveTeachers());
        stats.put("inactiveTeachers", teacherRepository.count() - teacherRepository.countActiveTeachers());
        stats.put("cachedTeachers", teachersCache.size());
        stats.put("lastImportTime", lastImportTime);
        stats.put("lastFileHash", lastFileHash);
        return stats;
    }
}