import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.FileOutputStream;

public class BarLineChart {

    public static void main(String[] args) throws Exception {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet chartdisplay=workbook.createSheet("ChartDisplay");
        Row row;
        Cell cell;

        row = chartdisplay.createRow(0);
        row.createCell(0);
        row.createCell(1).setCellValue("HEADER 1");
        row.createCell(2).setCellValue("HEADER 2");
        //row.createCell(3).setCellValue("HEADER 3");

        for (int r = 1; r < 8; r++) {
            row = chartdisplay.createRow(r);
            cell = row.createCell(0);
            cell.setCellValue("Serie " + r);
            cell = row.createCell(1);
            cell.setCellValue(new java.util.Random().nextDouble());
            cell = row.createCell(2);
            cell.setCellValue(new java.util.Random().nextDouble());
          //  cell = row.createCell(3);
            //cell.setCellValue(new java.util.Random().nextDouble());
        }

        XSSFDrawing drawing=chartdisplay.createDrawingPatriarch();
        ClientAnchor anchor=drawing.createAnchor(0,0,0,0,5,5,13,13);
        Chart chart=drawing.createChart(anchor);

        CTChart ctChart=((XSSFChart)chart).getCTChart();
        CTPlotArea ctPlotArea=ctChart.getPlotArea();
        //Bar Chart
        CTBarChart ctBarChart=ctPlotArea.addNewBarChart();
        CTBoolean ctBoolean=ctBarChart.addNewVaryColors();
        ctBoolean.setVal(false);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);
        CTBarSer ctBarSer=ctBarChart.addNewSer();
        CTSerTx ctSerTx=ctBarSer.addNewTx();
        CTStrRef ctStrRef=ctSerTx.addNewStrRef();
        ctStrRef.setF("\"BarSeriesName\"");
        //Labels For Bar Chart

        ctBarSer.addNewIdx().setVal(0); //0 = Color Blue
        CTAxDataSource ctAxDataSource=ctBarSer.addNewCat();
        ctStrRef=ctAxDataSource.addNewStrRef();
        String labelsRefer="ChartDisplay!A2:A7";//Excel Range where the Labels Are
        ctStrRef.setF(labelsRefer);
        //Values For Bar Chart
        CTNumDataSource ctNumDataSource=ctBarSer.addNewVal();
        CTNumRef ctNumRef=ctNumDataSource.addNewNumRef();
        String valuesRefer="ChartDisplay!B2:B7";//Excel range where values for barChart are
        ctNumRef.setF(valuesRefer);
        ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0,0,0});
        // Axis
        ctBarChart.addNewAxId().setVal(123456);
        ctBarChart.addNewAxId().setVal(123457);
        //cat axis
        CTCatAx ctCatAx=ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
        CTScaling ctScaling=ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.L);
        ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
        ctCatAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //val Left Axis
        CTValAx ctValAx1=ctPlotArea.addNewValAx();
        ctValAx1.addNewAxId().setVal(123457); //id of the val axis
        ctScaling=ctValAx1.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx1.addNewDelete().setVal(true);
        ctValAx1.addNewAxPos().setVal(STAxPos.L);
        ctValAx1.addNewMajorTickMark().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STTickMark.Enum.forInt(4));
        ctValAx1.addNewCrossAx().setVal(123456); //id of the cat axis
        ctValAx1.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctValAx1.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        ctValAx1.addNewMajorGridlines();

        // =======Line Chart
        //val Right Axis
        CTLineChart ctLineChart=ctPlotArea.addNewLineChart();
        CTBoolean ctBooleanLine=ctLineChart.addNewVaryColors();
        ctBooleanLine.setVal(false);
        CTLineSer ctLineSer=ctLineChart.addNewSer();
        CTSerTx ctSerTx1=ctLineSer.addNewTx();
        CTStrRef ctStrRef1=ctSerTx1.addNewStrRef();
        ctStrRef1.setF("\"LineSeriesName\"");
        ctLineSer.addNewIdx().setVal(2); //2= Color Grey
        CTAxDataSource ctAxDataSource1=ctLineSer.addNewCat();
        ctStrRef1=ctAxDataSource1.addNewStrRef();
        ctStrRef1.setF(labelsRefer);
        ctLineSer.addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0,0,0});

        String values2Refer="ChartDisplay!C2:C7"; //Excel Range Where Values for Line Values are
        CTNumDataSource ctNumDataSource1=ctLineSer.addNewVal();
        CTNumRef ctNumRef1=ctNumDataSource1.addNewNumRef();
        ctNumRef1.setF(values2Refer);

        //Axis
        ctLineChart.addNewAxId().setVal(1234);//id of the cat axis
        ctLineChart.addNewAxId().setVal(12345);

        CTCatAx ctCatAx1=ctPlotArea.addNewCatAx();
        ctCatAx1.addNewAxId().setVal(1234);// id of the cat Axis
        ctScaling=ctCatAx1.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx1.addNewDelete().setVal(true);
        ctCatAx1.addNewAxPos().setVal(STAxPos.R);
        ctCatAx1.addNewCrossAx().setVal(12345); //id of the val axis
        CTBoolean ctBoolean1=ctCatAx1.addNewAuto();


        CTValAx ctValAx=ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(12345); //id of the val axis
        ctScaling=ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.R);
        ctValAx.addNewCrossAx().setVal(1234); //id of the cat axis
        ctValAx.addNewMajorTickMark().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STTickMark.Enum.forInt(4));
        ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        //Legend
        CTLegend ctLegend=ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);

        System.out.println(ctChart);

        FileOutputStream fileOut = new FileOutputStream("BarLineChart.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }
}