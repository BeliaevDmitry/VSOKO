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
     * Создает Percentage Bar Chart на основе таблицы данных
     */
    public void createChartFromDataTable(XSSFWorkbook workbook, XSSFSheet sheet,
                                         int dataStartRow, int taskCount,
                                         int chartRow, String chartTitle) {

        try {
            // Создаем объект для рисования
            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // Определяем позицию графика
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
                    styleConfig.getLeftOffset(), chartRow,
                    styleConfig.getLeftOffset() + styleConfig.getColSpan(),
                    chartRow + styleConfig.getRowSpan());

            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText(chartTitle);

            // Легенда справа
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);

            // Оси
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("№ задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("% выполнения");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            // Формат для отображения процентов
            leftAxis.setNumberFormat("0%");

            // Данные для графика
            XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData barData = (XDDFBarChartData) data;
            barData.setBarDirection(BarDirection.COL);
            barData.setBarGrouping(BarGrouping.CLUSTERED);
            barData.setVaryColors(true);

            // Диапазоны данных
            CellRangeAddress labelRange = new CellRangeAddress(
                    dataStartRow + 1, dataStartRow + taskCount, 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            // Серия: % выполнения
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + taskCount, 5, 5));
            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle("% выполнения", null);

            // Рисуем график
            chart.plot(data);

            log.debug("✅ Percentage Bar Chart создан из таблицы данных");

        } catch (Exception e) {
            log.error("❌ Ошибка создания Percentage Bar Chart из таблицы данных: {}", e.getMessage(), e);
        }
    }
}