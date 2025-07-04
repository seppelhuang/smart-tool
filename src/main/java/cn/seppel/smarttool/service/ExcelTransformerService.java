package cn.seppel.smarttool.service;


import cn.seppel.smarttool.model.InvoiceRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelTransformerService {

    public void transform(String inputPath, String outputFile) throws Exception {

        File dir = new File(inputPath);
        List<InvoiceRow> rows = new ArrayList<>();
        File[] files = dir.listFiles();
        if (files == null || files.length == 0) {
            return;
        }
        for (File file : files) {

            String inputFile = file.getAbsolutePath();
            System.out.println("正在解析文件：" + inputFile);
            Map<String, String> containerMap = loadContainerMap(inputFile);

            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                for (Sheet sheet : workbook) {
                    if (sheet.getSheetName().equalsIgnoreCase("TOTAL")) continue;

                    String invoiceNo = getString(sheet.getRow(4).getCell(1)); // B5 = row 4, col 1

                    for (int i = 15; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        if (row == null || getString(row.getCell(0)).isEmpty()) continue;

                        InvoiceRow ir = new InvoiceRow();

                        ir.lineNo = getString(row.getCell(0));
                        if (!StringUtils.hasText(ir.lineNo) || ir.lineNo.endsWith("TOTAL")) {
                            break;
                        }
                        ir.poNo = getString(row.getCell(1));
                        ir.poLineNo = getString(row.getCell(2));
                        ir.itemNo = getString(row.getCell(3));
                        ir.description = getString(row.getCell(4));

                        String qty = getString(row.getCell(5));
                        int qtyInt = new BigDecimal(qty).intValue();
                        ir.qty = Integer.toString(qtyInt);
                        ir.unit = getString(row.getCell(6));
                        ir.price = getDouble(row.getCell(7));
                        ir.amount = getDouble(row.getCell(8));
                        ir.invoiceNo = invoiceNo;
                        String key = ir.poNo + "|" + ir.itemNo + "|" + qtyInt;
                        ir.container = containerMap.getOrDefault(key, "");

                        rows.add(ir);
                    }
                }
            }
        }


        writeOutput(rows, outputFile);
    }

    private Map<String, String> loadContainerMap(String file) throws Exception {
        Map<String, String> map = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {
            Sheet sheet = wb.getSheet("TOTAL");


            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String poNo = getString(row.getCell(1));
                String itemNo = getString(row.getCell(3));
                String qty = getString(row.getCell(5));
                String container = getString(row.getCell(11));
                String key = poNo + "|" + itemNo + "|" + qty;
                map.put(key, container);
            }
        }
        return map;
    }

    private void writeOutput(List<InvoiceRow> rows, String file) throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("output");
        String[] headers = {"Line No", "PO/NO.", "PO Line No", "Item NO", "Description", "Qty.", "Unit", "Price", "Amount", "Invoice No", "Container"};
        Row header = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) header.createCell(i).setCellValue(headers[i]);

        int r = 1;
        for (InvoiceRow row : rows) {
            Row out = sheet.createRow(r++);
            out.createCell(0).setCellValue(row.lineNo);
            out.createCell(1).setCellValue(row.poNo);
            out.createCell(2).setCellValue(row.poLineNo);
            out.createCell(3).setCellValue(row.itemNo);
            out.createCell(4).setCellValue(row.description);
            out.createCell(5).setCellValue(row.qty);
            out.createCell(6).setCellValue(row.unit);
            setNumericCell(out.createCell(7), row.price, wb, "0.00000");
            setNumericCell(out.createCell(8), row.amount, wb, "0.00");
            out.createCell(9).setCellValue(row.invoiceNo);
            out.createCell(10).setCellValue(row.container);
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            wb.write(fos);
        }
    }

    public static void setNumericCell(Cell cell, double value, Workbook workbook, String formatPattern) {
        CellStyle style = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat(formatPattern));
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static String getString(Cell cell) {
        if (cell == null) return "";
        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue().trim();
    }

    private static double getDouble(Cell cell) {
        if (cell == null) return 0;
        cell.setCellType(CellType.NUMERIC);
        return cell.getNumericCellValue();
    }
}
