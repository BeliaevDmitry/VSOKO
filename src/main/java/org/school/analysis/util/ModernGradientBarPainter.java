package org.school.analysis.util;

import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.ui.RectangleEdge;

import java.awt.*;
import java.awt.geom.RectangularShape;

/**
 * Современный градиентный painter для столбцов диаграмм.
 * Создает современный градиентный эффект для гистограмм.
 */
public class ModernGradientBarPainter extends StandardBarPainter {

    @Override
    public void paintBar(Graphics2D g2, BarRenderer renderer, int row, int column,
                         RectangularShape bar, RectangleEdge base) {

        // Получаем цвет для текущего столбца
        Color color = (Color) renderer.getItemPaint(row, column);

        if (color == null) {
            // Если цвет не установлен, используем стандартный
            super.paintBar(g2, renderer, row, column, bar, base);
            return;
        }

        // Сохраняем оригинальные настройки
        Paint originalPaint = g2.getPaint();
        Stroke originalStroke = g2.getStroke();

        try {
            // Создаем вертикальный градиент
            GradientPaint gradient;

            if (base == RectangleEdge.BOTTOM || base == RectangleEdge.TOP) {
                // Для горизонтальных столбцов
                gradient = new GradientPaint(
                        (float) bar.getMinX(), 0.0f,
                        color.brighter().brighter(),
                        (float) bar.getMaxX(), 0.0f,
                        color.darker()
                );
            } else {
                // Для вертикальных столбцов
                gradient = new GradientPaint(
                        0.0f, (float) bar.getMinY(),
                        color.brighter().brighter(),
                        0.0f, (float) bar.getMaxY(),
                        color.darker()
                );
            }

            // Применяем градиент
            g2.setPaint(gradient);
            g2.fill(bar);

            // Добавляем контур
            g2.setPaint(color.darker().darker());
            g2.setStroke(new BasicStroke(0.5f));
            g2.draw(bar);

        } finally {
            // Восстанавливаем оригинальные настройки
            g2.setPaint(originalPaint);
            g2.setStroke(originalStroke);
        }
    }
}