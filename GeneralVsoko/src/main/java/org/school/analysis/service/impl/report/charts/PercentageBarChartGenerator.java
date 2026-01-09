package org.school.analysis.service.impl.report.charts;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.school.analysis.model.dto.TaskStatisticsDto;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class PercentageBarChartGenerator extends ExcelChartBase {

    private final ChartStyleConfig styleConfig;

    @Autowired
    public PercentageBarChartGenerator(ChartStyleConfig styleConfig) {
        this.styleConfig = styleConfig;
    }

    /**
     * Создает Bar Chart для процента выполнения
     */
    public void createChart(XSSFWorkbook workbook, XSSFSheet sheet,
                            List<TaskStatisticsDto> tasks,
                            int dataStartRow, int chartRow,
                            String chartTitle) {

        try {
            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // Позиция графика
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    styleConfig.getLeftOffset(), chartRow + 1,
                    styleConfig.getLeftOffset() + styleConfig.getColSpan(),
                    chartRow + styleConfig.getRowSpan());

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText(chartTitle);

            // Легенда справа
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);

            // Оси
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Задание");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("% выполнения");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            leftAxis.setMinimum(0.0);
            leftAxis.setMaximum(1.0);

            // Данные
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.STANDARD);
            barData.setVaryColors(true);

            // Диапазоны данных
            CellRangeAddress labelRange = new CellRangeAddress(
                    dataStartRow + 1, dataStartRow + tasks.size(), 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + tasks.size(), 4, 4));

            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle("% выполнения", null);

            chart.plot(data);

            log.debug("✅ Percentage Bar Chart создан");

        } catch (Exception e) {
            log.error("❌ Ошибка создания Percentage Bar Chart: {}", e.getMessage(), e);
        }
    }
}