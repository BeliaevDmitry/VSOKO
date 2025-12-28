package org.school.analysis.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Утилиты для работы с файлами и директориями
 */
public class FileHelper {

    /**
     * Получение всех Excel файлов из папки
     */
    public static List<File> getExcelFiles(String folderPath) {
        return getFilesByExtension(folderPath, ".xlsx");
    }

    /**
     * Получение файлов по расширению
     */
    public static List<File> getFilesByExtension(String folderPath, String extension) {
        List<File> files = new ArrayList<>();
        File folder = new File(folderPath);

        if (!folder.exists() || !folder.isDirectory()) {
            throw new IllegalArgumentException("Папка не существует: " + folderPath);
        }

        File[] foundFiles = folder.listFiles((dir, name) ->
                name.toLowerCase().endsWith(extension.toLowerCase())
        );

        if (foundFiles != null) {
            for (File file : foundFiles) {
                if (file.isFile()) {
                    files.add(file);
                }
            }
        }

        return files;
    }

    /**
     * Создание директории, если её нет
     */
    public static void createDirectoryIfNotExists(String dirPath) throws IOException {
        Path path = Paths.get(dirPath);
        if (!Files.exists(path)) {
            Files.createDirectories(path);
        }
    }

    /**
     * Копирование файла
     */
    public static void copyFile(File source, File destination) throws IOException {
        Files.copy(source.toPath(), destination.toPath(),
                StandardCopyOption.REPLACE_EXISTING);
    }

    /**
     * Перемещение файла
     */
    public static void moveFile(File source, File destination) throws IOException {
        Files.move(source.toPath(), destination.toPath(),
                StandardCopyOption.REPLACE_EXISTING);
    }

    /**
     * Удаление файла
     */
    public static boolean deleteFile(File file) {
        try {
            return Files.deleteIfExists(file.toPath());
        } catch (IOException e) {
            return false;
        }
    }

    /**
     * Рекурсивное удаление директории
     */
    public static void deleteDirectory(File directory) throws IOException {
        if (!directory.exists()) {
            return;
        }

        Files.walkFileTree(directory.toPath(), new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs)
                    throws IOException {
                Files.delete(file);
                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult postVisitDirectory(Path dir, IOException exc)
                    throws IOException {
                Files.delete(dir);
                return FileVisitResult.CONTINUE;
            }
        });
    }

    /**
     * Получение размера директории (в байтах)
     */
    public static long getDirectorySize(File directory) {
        if (!directory.exists() || !directory.isDirectory()) {
            return 0;
        }

        long size = 0;
        File[] files = directory.listFiles();

        if (files != null) {
            for (File file : files) {
                if (file.isFile()) {
                    size += file.length();
                } else {
                    size += getDirectorySize(file);
                }
            }
        }

        return size;
    }

    /**
     * Получение человеко-читаемого размера
     */
    public static String getHumanReadableSize(long bytes) {
        if (bytes < 1024) {
            return bytes + " B";
        }

        int exp = (int) (Math.log(bytes) / Math.log(1024));
        String pre = "KMGTPE".charAt(exp - 1) + "";
        return String.format("%.1f %sB", bytes / Math.pow(1024, exp), pre);
    }

    /**
     * Проверка, является ли файл архивом
     */
    public static boolean isZipFile(File file) {
        if (file == null || !file.exists() || !file.isFile()) {
            return false;
        }

        try (ZipInputStream zis = new ZipInputStream(new java.io.FileInputStream(file))) {
            ZipEntry entry = zis.getNextEntry();
            return entry != null;
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * Извлечение расширения файла
     */
    public static String getFileExtension(File file) {
        String name = file.getName();
        int lastDot = name.lastIndexOf('.');
        return lastDot > 0 ? name.substring(lastDot).toLowerCase() : "";
    }

    /**
     * Получение имени файла без расширения
     */
    public static String getFileNameWithoutExtension(File file) {
        String name = file.getName();
        int lastDot = name.lastIndexOf('.');
        return lastDot > 0 ? name.substring(0, lastDot) : name;
    }

    /**
     * Проверка, содержит ли имя файла паттерн
     */
    public static boolean fileNameContains(File file, String pattern) {
        return file.getName().toLowerCase().contains(pattern.toLowerCase());
    }

    /**
     * Создание временного файла
     */
    public static File createTempFile(String prefix, String suffix) throws IOException {
        return File.createTempFile(prefix, suffix);
    }

    /**
     * Запись текста в файл
     */
    public static void writeTextToFile(String text, File file) throws IOException {
        Files.writeString(file.toPath(), text);
    }

    /**
     * Чтение текста из файла
     */
    public static String readTextFromFile(File file) throws IOException {
        return Files.readString(file.toPath());
    }
}