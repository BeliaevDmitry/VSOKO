package org.school.analysis.util;

import java.awt.*;

/**
 * Утилиты для создания графиков и диаграмм
 */
public class ChartStyleUtils {

    private ChartStyleUtils() {
        // Utility class
    }

    /**
     * Возвращает цвет для процента выполнения (от красного к зеленому)
     */
    public static Color getColorForPercentage(double percentage) {
        if (percentage >= 90) return new Color(39, 174, 96);    // Зеленый
        if (percentage >= 70) return new Color(46, 204, 113);   // Светло-зеленый
        if (percentage >= 50) return new Color(241, 196, 15);   // Желтый
        if (percentage >= 30) return new Color(230, 126, 34);   // Оранжевый
        return new Color(231, 76, 60);                         // Красный
    }

    /**
     * Возвращает цвет для тепловой карты
     */
    public static Color getHeatMapColor(double percentage) {
        // От темно-синего (сложное) к светло-зеленому (легкое)
        if (percentage >= 90) return new Color(162, 222, 150);  // Светло-зеленый
        if (percentage >= 70) return new Color(117, 199, 111);  // Зеленый
        if (percentage >= 50) return new Color(72, 176, 72);    // Темно-зеленый
        if (percentage >= 30) return new Color(44, 127, 184);   // Синий
        return new Color(37, 52, 148);                         // Темно-синий
    }

    /**
     * Создает градиент между двумя цветами
     */
    public static Color getGradientColor(Color start, Color end, double ratio) {
        int red = (int) (start.getRed() * (1 - ratio) + end.getRed() * ratio);
        int green = (int) (start.getGreen() * (1 - ratio) + end.getGreen() * ratio);
        int blue = (int) (start.getBlue() * (1 - ratio) + end.getBlue() * ratio);

        red = Math.min(Math.max(red, 0), 255);
        green = Math.min(Math.max(green, 0), 255);
        blue = Math.min(Math.max(blue, 0), 255);

        return new Color(red, green, blue);
    }

    /**
     * Расчет линейного тренда методом наименьших квадратов
     */
    public static double[] calculateLinearTrend(java.util.List<Double> values) {
        if (values == null || values.size() < 2) {
            return new double[]{0, 0};
        }

        int n = values.size();
        double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

        for (int i = 0; i < n; i++) {
            double x = i + 1;
            double y = values.get(i);
            sumX += x;
            sumY += y;
            sumXY += x * y;
            sumX2 += x * x;
        }

        double slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        double intercept = (sumY - slope * sumX) / n;

        return new double[]{slope, intercept};
    }

    /**
     * Создает массив цветов для категорий
     */
    public static Color[] createModernColorPalette(int count) {
        Color[] palette = new Color[count];

        // Современная цветовая палитра
        Color[] baseColors = {
                new Color(52, 152, 219),    // Ярко-синий
                new Color(46, 204, 113),    // Ярко-зеленый
                new Color(155, 89, 182),    // Фиолетовый
                new Color(241, 196, 15),    // Желтый
                new Color(230, 126, 34),    // Оранжевый
                new Color(231, 76, 60),     // Красный
                new Color(149, 165, 166),   // Серый
                new Color(26, 188, 156)     // Бирюзовый
        };

        for (int i = 0; i < count; i++) {
            palette[i] = baseColors[i % baseColors.length];
        }

        return palette;
    }
}