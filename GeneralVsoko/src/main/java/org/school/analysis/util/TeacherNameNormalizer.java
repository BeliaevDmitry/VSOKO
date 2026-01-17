package org.school.analysis.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.text.similarity.LevenshteinDistance;

import java.text.Normalizer;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Slf4j
public class TeacherNameNormalizer {

    private static final Pattern DIACRITICS = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
    private static final Pattern MULTIPLE_SPACES = Pattern.compile("\\s+");
    private static final Pattern PUNCTUATION = Pattern.compile("[.,!?;:()\\[\\]{}<>\\/\\\\\"'`~@#$%^&*_+=|]");
    private static final Pattern RUSSIAN_LETTERS = Pattern.compile("[а-яА-ЯёЁ\\s-]+");
    private static final LevenshteinDistance LEVENSHTEIN = new LevenshteinDistance();

    // Регулярное выражение для инициалов с точками и пробелами
    private static final Pattern INITIALS_PATTERN =
            Pattern.compile("([а-я])\\.?\\s*([а-я])\\.?\\s*([а-я])?\\.?");

    // Типичные сокращения
    private static final Map<String, String> COMMON_SHORTENINGS = Stream.of(
            new String[][] {
                    {"алл.", "алла"},
                    {"андр.", "андрей"},
                    {"анн.", "анна"},
                    {"влад.", "владимир"},
                    {"дмитр.", "дмитрий"},
                    {"евг.", "евгений"},
                    {"макс.", "максим"},
                    {"мих.", "михаил"},
                    {"ник.", "николай"},
                    {"олег", "олег"},
                    {"петр", "петр"},
                    {"серг.", "сергей"},
                    {"тат.", "татьяна"},
                    {"юл.", "юлия"},
                    {"юр.", "юрий"}
            }).collect(Collectors.toMap(data -> data[0], data -> data[1]));

    /**
     * Нормализация имени для сравнения
     */
    public static String normalize(String name) {
        if (name == null || name.trim().isEmpty()) {
            return "";
        }

        String result = name.toLowerCase();

        // 1. Удаляем пунктуацию
        result = PUNCTUATION.matcher(result).replaceAll("");

        // 2. Заменяем все последовательности пробелов одним пробелом
        result = MULTIPLE_SPACES.matcher(result.trim()).replaceAll(" ");

        // 3. Нормализуем русские символы
        result = normalizeRussian(result);

        // 4. Обработка инициалов (например, "и и" -> "и и")
        result = normalizeInitials(result);

        // 5. Убираем диакритические знаки (если есть)
        result = DIACRITICS.matcher(Normalizer.normalize(result, Normalizer.Form.NFD))
                .replaceAll("");

        // 6. Заменяем сокращения на полные формы
        result = expandShortenings(result);

        // 7. Сортируем части ФИО
        result = sortNameParts(result);

        return result.trim();
    }

    /**
     * Нормализация русских символов
     */
    private static String normalizeRussian(String text) {
        return text.toLowerCase()
                .replace('a', 'а')
                .replace('b', 'в')
                .replace('c', 'с')
                .replace('e', 'е')
                .replace('h', 'н')
                .replace('k', 'к')
                .replace('m', 'м')
                .replace('o', 'о')
                .replace('p', 'р')
                .replace('t', 'т')
                .replace('x', 'х')
                .replace('y', 'у');
    }

    /**
     * Нормализация инициалов
     * Примеры:
     * - "и.и." → "и и"
     * - "и . и ." → "и и"
     * - "и  и  " → "и и"
     * - "иванов и и" → "иванов и и"
     */
    private static String normalizeInitials(String text) {
        // Удаляем точки возле букв
        String result = text.replaceAll("([а-я])\\.", "$1");

        // Если текст похож на инициалы (1-2 буквы с пробелами)
        String[] parts = result.split("\\s+");
        if (parts.length >= 2) {
            // Проверяем, все ли части - одиночные буквы
            boolean allAreSingleLetters = true;
            for (String part : parts) {
                if (part.length() != 1 || !Character.isLetter(part.charAt(0))) {
                    allAreSingleLetters = false;
                    break;
                }
            }

            // Если все части - одиночные буквы, это инициалы
            if (allAreSingleLetters && parts.length <= 3) {
                return String.join(" ", parts);
            }
        }

        return result;
    }

    /**
     * Расширение сокращений
     */
    private static String expandShortenings(String text) {
        String result = text;
        for (Map.Entry<String, String> entry : COMMON_SHORTENINGS.entrySet()) {
            // Используем границы слов для точного совпадения
            result = result.replaceAll("\\b" + Pattern.quote(entry.getKey()) + "\\b", entry.getValue());
        }
        return result;
    }

    /**
     * Сортировка частей имени для унификации
     */
    private static String sortNameParts(String name) {
        String[] parts = name.split("\\s+");
        if (parts.length <= 1) {
            return name;
        }

        // Ищем фамилию (обычно самая длинная часть или первая)
        String likelySurname = findLikelySurname(parts);

        List<String> otherParts = new ArrayList<>();
        for (String part : parts) {
            if (!part.equals(likelySurname)) {
                otherParts.add(part);
            }
        }

        // Собираем: фамилия + остальные части
        List<String> result = new ArrayList<>();
        result.add(likelySurname);
        result.addAll(otherParts);

        return String.join(" ", result);
    }

    /**
     * Поиск наиболее вероятной фамилии
     */
    private static String findLikelySurname(String[] parts) {
        // Если первая часть длинная (>3 символов), это скорее всего фамилия
        if (parts[0].length() > 3) {
            return parts[0];
        }

        // Ищем самую длинную часть
        String longestPart = parts[0];
        for (String part : parts) {
            if (part.length() > longestPart.length()) {
                longestPart = part;
            }
        }

        return longestPart;
    }

    /**
     * Получение короткого имени (Фамилия И.О.)
     */
    public static String getShortName(String fullName) {
        String normalized = normalize(fullName);
        String[] parts = normalized.split("\\s+");

        if (parts.length < 2) {
            return capitalizeName(fullName);
        }

        StringBuilder shortName = new StringBuilder(capitalizeFirst(parts[0])); // Фамилия

        for (int i = 1; i < Math.min(parts.length, 3); i++) {
            if (!parts[i].isEmpty()) {
                shortName.append(" ")
                        .append(Character.toUpperCase(parts[i].charAt(0)))
                        .append(".");
            }
        }

        return shortName.toString();
    }

    /**
     * Проверка схожести имен
     */
    public static boolean isSimilar(String name1, String name2) {
        if (name1 == null || name2 == null) {
            return false;
        }

        String norm1 = normalize(name1);
        String norm2 = normalize(name2);

        // Быстрое сравнение
        if (norm1.equals(norm2)) {
            return true;
        }

        // Проверка, содержатся ли части одного имени в другом
        if (containsNameParts(norm1, norm2)) {
            return true;
        }

        // Проверка по Левенштейну
        Integer distance = LEVENSHTEIN.apply(norm1, norm2);
        int maxLength = Math.max(norm1.length(), norm2.length());

        if (maxLength == 0) {
            return false;
        }

        double similarity = 1.0 - (double) distance / maxLength;
        return similarity >= 0.7; // 70% похожести достаточно
    }

    /**
     * Проверяет, содержатся ли части одного имени в другом
     */
    private static boolean containsNameParts(String name1, String name2) {
        Set<String> parts1 = new HashSet<>(Arrays.asList(name1.split("\\s+")));
        Set<String> parts2 = new HashSet<>(Arrays.asList(name2.split("\\s+")));

        // Находим пересечение
        parts1.retainAll(parts2);
        return !parts1.isEmpty();
    }

    /**
     * Приведение имени к нормальному виду (с заглавными буквами)
     */
    private static String capitalizeName(String name) {
        if (name == null || name.isEmpty()) {
            return name;
        }

        String[] parts = name.split("\\s+");
        StringBuilder result = new StringBuilder();

        for (String part : parts) {
            if (!part.isEmpty()) {
                if (result.length() > 0) {
                    result.append(" ");
                }

                if (part.length() == 1) {
                    // Одиночная буква (инициал)
                    result.append(part.toUpperCase());
                } else {
                    // Полные части имени
                    result.append(Character.toUpperCase(part.charAt(0)))
                            .append(part.substring(1));
                }
            }
        }

        return result.toString();
    }

    /**
     * Извлечение фамилии из строки
     */
    public static String extractLastName(String name) {
        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");

        if (parts.length == 0) {
            return "";
        }

        // Ищем фамилию (первая часть или самая длинная)
        String lastName = parts[0];
        for (String part : parts) {
            if (part.length() > lastName.length() && part.length() > 2) {
                lastName = part;
            }
        }

        return capitalizeFirst(lastName);
    }

    /**
     * Извлечение инициалов из строки
     */
    public static String extractInitials(String name) {
        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");

        StringBuilder initials = new StringBuilder();
        for (String part : parts) {
            if (part.length() == 1 && Character.isLetter(part.charAt(0))) {
                initials.append(part.toUpperCase()).append(".");
            }
        }

        return initials.toString();
    }

    private static String capitalizeFirst(String word) {
        if (word == null || word.isEmpty()) {
            return word;
        }
        return Character.toUpperCase(word.charAt(0)) + word.substring(1);
    }

    /**
     * Дополнительный метод для строгой проверки формата "Фамилия И.О."
     */
    public static boolean isValidShortNameFormat(String name) {
        if (name == null) {
            return false;
        }

        String normalized = normalize(name);
        String[] parts = normalized.split("\\s+");

        // Должно быть как минимум 2 части (фамилия и один инициал)
        if (parts.length < 2) {
            return false;
        }

        // Первая часть (фамилия) должна быть длиннее 1 символа
        if (parts[0].length() <= 1) {
            return false;
        }

        // Остальные части должны быть одиночными буквами (инициалы)
        for (int i = 1; i < parts.length; i++) {
            if (parts[i].length() != 1 || !Character.isLetter(parts[i].charAt(0))) {
                return false;
            }
        }

        return true;
    }
}