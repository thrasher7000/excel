package com.mycompany.reporteexcel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.FileOutputStream;
import java.math.BigDecimal;

public class Main {

    public static void main(String[] args) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(0, "Hoja excel");
        

        String[] headers = new String[]{
            "Placa",
            "Precio Mantenimientoo",
            "Tipo vehículo"
        };

        Object[][] data = new Object[][] {
            new Object[] { "HFG-852", new BigDecimal("340.95"), "Camión" },
            new Object[] { "POF-890", new BigDecimal("41.95"), "Carro" },
            new Object[] { "OHG-752 ", new BigDecimal("421.36"), "Furgoneta" }
        };

        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerStyle.setFont(font);

        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        HSSFRow headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            HSSFCell cell = headerRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header);
        }

        for (int i = 0; i < data.length; ++i) {
            HSSFRow dataRow = sheet.createRow(i + 1);

            Object[] d = data[i];
            String product = (String) d[0];
            BigDecimal price = (BigDecimal) d[1];
            String link = (String) d[2];

            dataRow.createCell(0).setCellValue(product);
            dataRow.createCell(1).setCellValue(price.doubleValue());
            dataRow.createCell(2).setCellValue(link);
        }

        HSSFRow dataRow = sheet.createRow(1 + data.length);
        HSSFCell total = dataRow.createCell(1);
        total.setCellType(Cell.CELL_TYPE_FORMULA);
        total.setCellStyle(style);
        total.setCellFormula(String.format("SUM(B2:B%d)", 1 + data.length));
        
        
        
        FileOutputStream file = new FileOutputStream("C:\\\\Users\\\\miguelangel\\\\Documents\\\\NetBeansProjects\\\\reporte\\\\workbook.xls");
        workbook.write(file);
        file.close();
    }
}
