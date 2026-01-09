package org.school.analysis.service.impl.report.charts;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series;

import java.awt.Color;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;

/**
 * БАЗОВЫЙ КЛАСС ДЛЯ ГЕНЕРАТОРОВ ГРАФИКОВ
 * Работает с POI 5.3.0
 */
@Slf4j
public abstract class ExcelChartBase {

    /**
     * Устанавливает цвет для серии диаграммы
     */
    protected void setSeriesColor(XSSFChart chart, int seriesIndex, Color color) {
        try {
            if (chart.getChartSeries() == null || chart.getChartSeries().isEmpty()) {
                return;
            }

            XDDFChartData data = chart.getChartSeries().get(0);
            if (seriesIndex >= data.getSeriesCount()) {
                return;
            }

            Series series = data.getSeries(seriesIndex);
            XDDFShapeProperties shapeProps = series.getShapeProperties();
            if (shapeProps == null) {
                shapeProps = new XDDFShapeProperties();
            }

            // Создаем цвет для POI 5.3.0
            XDDFSolidFillProperties fill = createSolidFill(color);
            shapeProps.setFillProperties(fill);
            series.setShapeProperties(shapeProps);

        } catch (Exception e) {
            log.warn("⚠️ Не удалось установить цвет для серии {}: {}", seriesIndex, e.getMessage());
        }
    }

    /**
     * Создает SolidFillProperties для цвета
     */
    protected XDDFSolidFillProperties createSolidFill(Color color) {
        try {
            XDDFColor xddfColor = createXDDFColor(color);
            return new XDDFSolidFillProperties(xddfColor);
        } catch (Exception e) {
            log.warn("Не удалось создать цвет, используем синий по умолчанию: {}", e.getMessage());
            return createDefaultSolidFill();
        }
    }

    /**
     * Создает цвет по умолчанию
     */
    private XDDFSolidFillProperties createDefaultSolidFill() {
        try {
            // Попробуем разные способы
            XDDFColor defaultColor = createXDDFColor(Color.BLUE);
            return new XDDFSolidFillProperties(defaultColor);
        } catch (Exception e) {
            // Последний резервный вариант
            return new XDDFSolidFillProperties((XDDFColor) null);
        }
    }

    /**
     * Создает XDDFColor из java.awt.Color
     */
    protected XDDFColor createXDDFColor(Color color) throws Exception {
        // Пробуем разные способы по порядку
        try {
            return createXDDFColorMethod1(color);
        } catch (Exception e1) {
            try {
                return createXDDFColorMethod2(color);
            } catch (Exception e2) {
                try {
                    return createXDDFColorMethod3(color);
                } catch (Exception e3) {
                    throw new Exception("Не удалось создать XDDFColor");
                }
            }
        }
    }

    /**
     * Способ 1: Через XDDFColorRgbBinary
     */
    private XDDFColor createXDDFColorMethod1(Color color) throws Exception {
        try {
            Class<?> rgbBinaryClass = Class.forName("org.apache.poi.xddf.usermodel.XDDFColorRgbBinary");
            byte[] rgb = new byte[3];
            rgb[0] = (byte) color.getRed();
            rgb[1] = (byte) color.getGreen();
            rgb[2] = (byte) color.getBlue();

            Constructor<?> constructor = rgbBinaryClass.getConstructor(byte[].class);
            return (XDDFColor) constructor.newInstance((Object) rgb);
        } catch (Exception e) {
            throw new Exception("Способ 1 не сработал: " + e.getMessage());
        }
    }

    /**
     * Способ 2: Через рефлексию и метод from
     */
    private XDDFColor createXDDFColorMethod2(Color color) throws Exception {
        try {
            // Ищем статический метод from
            Method fromMethod = XDDFColor.class.getMethod("from", int.class);
            return (XDDFColor) fromMethod.invoke(null, color.getRGB());
        } catch (NoSuchMethodException e) {
            try {
                // Пробуем другой вариант метода from
                Method fromMethod = XDDFColor.class.getMethod("from", String.class);
                String hex = String.format("%06x", color.getRGB() & 0xFFFFFF);
                return (XDDFColor) fromMethod.invoke(null, hex);
            } catch (NoSuchMethodException e2) {
                throw new Exception("Способ 2 не сработал");
            }
        }
    }

    /**
     * Способ 3: Через создание конкретной реализации
     */
    private XDDFColor createXDDFColorMethod3(Color color) throws Exception {
        try {
            // Пытаемся создать через XDDFColorRgbPercent
            Class<?> rgbPercentClass = Class.forName("org.apache.poi.xddf.usermodel.XDDFColorRgbPercent");

            // Создаем массив процентов (0-100)
            int[] percent = new int[3];
            percent[0] = (int) (color.getRed() / 255.0 * 100);
            percent[1] = (int) (color.getGreen() / 255.0 * 100);
            percent[2] = (int) (color.getBlue() / 255.0 * 100);

            Constructor<?> constructor = rgbPercentClass.getConstructor(int.class, int.class, int.class);
            return (XDDFColor) constructor.newInstance(percent[0], percent[1], percent[2]);

        } catch (Exception e) {
            throw new Exception("Способ 3 не сработал: " + e.getMessage());
        }
    }
}