package org.school.analysis.service.impl.report.charts;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.awt.Color;
import java.util.List;

@Component
@Slf4j
public class StackedBarChartGenerator extends ExcelChartBase {

    private final ChartStyleConfig styleConfig;

    @Autowired
    public StackedBarChartGenerator(ChartStyleConfig styleConfig) {
        this.styleConfig = styleConfig;
    }

    /**
     * Создает Stacked Bar Chart
     */
    public void createChart(XSSFWorkbook workbook, XSSFSheet sheet,
                            List<TaskStatisticsDto> tasks,
                            int dataStartRow, int chartRow,
                            String chartTitle) {

        try {
            // Создаем объект для рисования
            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // Определяем позицию графика
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    styleConfig.getLeftOffset(), chartRow + 1,
                    styleConfig.getLeftOffset() + styleConfig.getColSpan(),
                    chartRow + styleConfig.getRowSpan());

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText(chartTitle);

            // Легенда внизу
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.BOTTOM);

            // Оси
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("№ задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Количество студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            // Данные для графика
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STACKED);
            barData.setVaryColors(true);

            // Диапазоны данных
            CellRangeAddress labelRange = new CellRangeAddress(
                    dataStartRow + 1, dataStartRow + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            // Серия 1: Полностью выполнившие
            XDDFNumericalDataSource<Double> ys1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 1, 1));
            XDDFChartData.Series series1 = data.addSeries(xs, ys1);
            series1.setTitle("Полностью", null);

            // Серия 2: Частично выполнившие
            XDDFNumericalDataSource<Double> ys2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 2, 2));
            XDDFChartData.Series series2 = data.addSeries(xs, ys2);
            series2.setTitle("Частично", null);

            // Серия 3: Не выполнившие
            XDDFNumericalDataSource<Double> ys3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 3, 3));
            XDDFChartData.Series series3 = data.addSeries(xs, ys3);
            series3.setTitle("Не справилось", null);

            // Рисуем график
            chart.plot(data);

            log.debug("✅ Stacked Bar Chart создан");

        } catch (Exception e) {
            log.error("❌ Ошибка создания Stacked Bar Chart: {}", e.getMessage(), e);
        }
    }
}