package org.school.analysis.service.impl.report.charts;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetLineDash;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.awt.Color;
import java.lang.reflect.Method;

@Component
@Slf4j
public class LineChartGenerator extends ExcelChartBase {

    private final ChartStyleConfig styleConfig;

    @Autowired
    public LineChartGenerator(ChartStyleConfig styleConfig) {
        this.styleConfig = styleConfig;
    }

    /**
     * Создает Line Chart (линейный график) на основе таблицы данных
     */
    public void createChartFromDataTable(XSSFWorkbook workbook, XSSFSheet sheet,
                                         int dataStartRow, int taskCount,
                                         int chartRow, String chartTitle) {

        try {
            // 1. Безопасное получение drawing
            XSSFDrawing drawing = sheet.getDrawingPatriarch();
            if (drawing == null) {
                drawing = sheet.createDrawingPatriarch();
            }

            // 2. Расчет позиции
            int startCol = styleConfig.getLeftOffset();
            int startRow = chartRow;
            int endCol = startCol + styleConfig.getColSpan();
            int endRow = startRow + styleConfig.getRowSpan();

            // 3. Создание якоря
            XSSFClientAnchor anchor = drawing.createAnchor(
                    0, 0, 0, 0,
                    startCol, startRow,
                    endCol, endRow
            );

            // 4. Создание графика
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText(chartTitle);

            // 5. Упростить создание данных (убрать reflection для теста)
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Номер задания");

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("% выполнения");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            // 6. Простое создание данных БЕЗ сложных настроек
            XDDFChartData data = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

            CellRangeAddress labelRange = new CellRangeAddress(
                    dataStartRow + 1, dataStartRow + taskCount, 0, 0);
            XDDFDataSource<String> xs = XDDFDataSourcesFactory.fromStringCellRange(sheet, labelRange);

            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
                    new CellRangeAddress(dataStartRow + 1, dataStartRow + taskCount, 5, 5));

            XDDFChartData.Series series = data.addSeries(xs, ys);
            series.setTitle("% выполнения", null);

            // 7. Простой plot без дополнительных настроек
            chart.plot(data);

            log.debug("✅ Line Chart создан из таблицы данных");

        } catch (Exception e) {
            log.error("❌ Ошибка создания Line Chart: {}", e.getMessage(), e);
            // Не выбрасывайте исключение дальше, чтобы не ломать весь отчет
        }
    }
}