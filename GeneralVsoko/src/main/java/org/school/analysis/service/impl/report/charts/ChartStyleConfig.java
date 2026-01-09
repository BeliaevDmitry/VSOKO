package org.school.analysis.service.impl.report.charts;

import lombok.Getter;
import lombok.Setter;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.awt.*;

/**
 * КОНФИГУРАЦИЯ СТИЛЕЙ ГРАФИКОВ
 * Все параметры можно настраивать через application.yml
 */
@Component
@ConfigurationProperties(prefix = "app.excel.charts")
@Getter
@Setter
public class ChartStyleConfig {

    // ============ РАЗМЕРЫ И ПОЗИЦИИ ============

    /** Ширина графика в колонках */
    private int colSpan = 10;

    /** Высота графика в строках */
    private int rowSpan = 15;

    /** Отступ между графиками в строках */
    private int spacing = 5;

    /** Отступ графика от левого края в колонках */
    private int leftOffset = 0;

    // ============ ЦВЕТА ГРАФИКОВ ============

    /** Цвет для "Полностью выполнивших" в формате RGB (R,G,B) */
    private String fullyCompletedColor = "0,20,236";

    /** Цвет для "Частично выполнивших" */
    private String partiallyCompletedColor = "241,120,0";

    /** Цвет для "Не выполнивших" */
    private String notCompletedColor = "231,50,60";

    /** Цвет для графика процента выполнения */
    private String percentageColor = "52,152,219";

    /** Цвет линии на линейном графике */
    private String lineColor = "52,152,219";

    // ============ СТИЛИ ЛИНИЙ ============

    /** Толщина линии на линейном графике */
    private double lineWidth = 2.0;

    /** Размер маркера на линейном графике */
    private short markerSize = 6;

    // ============ МЕТОДЫ ДЛЯ ПРЕОБРАЗОВАНИЯ ЦВЕТОВ ============

    public Color getFullyCompletedColor() {
        return parseColor(fullyCompletedColor);
    }

    public Color getPartiallyCompletedColor() {
        return parseColor(partiallyCompletedColor);
    }

    public Color getNotCompletedColor() {
        return parseColor(notCompletedColor);
    }

    public Color getPercentageColor() {
        return parseColor(percentageColor);
    }

    public Color getLineColor() {
        return parseColor(lineColor);
    }

    private Color parseColor(String rgbString) {
        try {
            String[] parts = rgbString.split(",");
            if (parts.length == 3) {
                int r = Integer.parseInt(parts[0].trim());
                int g = Integer.parseInt(parts[1].trim());
                int b = Integer.parseInt(parts[2].trim());
                return new Color(r, g, b);
            }
        } catch (Exception e) {
            // Ошибка парсинга - возвращаем цвет по умолчанию
        }
        return Color.BLUE;
    }
}