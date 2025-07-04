
package cn.seppel.smarttool.pdfexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PdfToExcelWithHeader {

    public static void main(String[] args) throws IOException {
        String inputPdf = "D:\\2.pdf";
        String outputExcel = "D:\\output_with_header.xlsx";

        String content;
        try (PDDocument document = PDDocument.load(new File(inputPdf))) {
            PDFTextStripper stripper = new PDFTextStripper();
            content = stripper.getText(document);
        }

        List<LineItem> items = extractLineItemsFromText(content);
        Map<String, String> headerInfo = extractHeaderInfo(content);

        exportWithHeaderAndItems(headerInfo, items, outputExcel);

        System.out.println("导出成功（包含头部信息）: " + outputExcel);
    }

    public static List<LineItem> extractLineItemsFromText(String content) {
        List<LineItem> list = new ArrayList<>();
        String[] lines = content.split("\r?\n");
        LineItem current = null;

        for (String line : lines) {
            line = line.trim();

            if (line.matches("^\\d{2}\\s+\\d+\\s+EA\\s+Delta Item\\s*:\\s*.*")) {
                if (current != null) list.add(current);
                current = new LineItem();

                Pattern pattern = Pattern.compile("^(\\d{2})\\s+(\\d+)\\s+(EA)\\s+Delta Item\\s*:\\s*(\\s+)");
                Matcher matcher = pattern.matcher(line);
                if (matcher.find()) {
                    current.setLineNo(matcher.group(1));
                    current.setQuantity(matcher.group(2));
                    current.setUnit(matcher.group(3));
                    current.setDeltaItem(matcher.group(4));
                }
            } else if (line.startsWith("Description:")) {
                current.setDescription(line.replace("Description:", "").trim());
            } else if (line.startsWith("Rev Level:")) {
                current.setRevLevel(line.replace("Rev Level:", "").trim());
            } else if (line.startsWith("Contractual Delivery Date:")) {
                current.setDeliveryDate(line.replace("Contractual Delivery Date:", "").trim());
            } else if (line.matches("^\\d{1,3}(,\\d{3})*(\\.\\d{2})?\\s+\\d{1,3}(,\\d{3})*(\\.\\d{2})?$")) {
                String[] parts = line.split("\\s+");
                if (parts.length == 2) {
                    current.setUnitPrice(parts[0]);
                    current.setExtendedPrice(parts[1]);
                }
            }
        }
        if (current != null) list.add(current);
        return list;
    }

    public static Map<String, String> extractHeaderInfo(String content) {
        Map<String, String> info = new LinkedHashMap<>();

        Pattern poPattern = Pattern.compile("PO\\s*:\\s*(\\d+)");
        Pattern vendorPattern = Pattern.compile("Vendor:\\s*(\\d+)");
        Pattern plantPattern = Pattern.compile("Plant\\s*:(\\d+)");
        Pattern datePattern = Pattern.compile("Purchase Order Date[:：]\\s*(\\d{1,2}/\\d{1,2}/\\d{2})");
        Pattern buyerPattern = Pattern.compile("Buyer Name:\\s*(.+)");
        Pattern emailPattern = Pattern.compile("EMail:\\s*(\\s+@\\s+)");
        Pattern supplierEmailPattern = Pattern.compile("Supplier Contact:\\s*\\s*([\\w.,@\\-]+)");

        info.put("PO", findFirstMatch(poPattern, content));
        info.put("Vendor", findFirstMatch(vendorPattern, content));
        info.put("Plant", findFirstMatch(plantPattern, content));
        info.put("Purchase Date", findFirstMatch(datePattern, content));
        info.put("Buyer Name", findFirstMatch(buyerPattern, content));
        info.put("Buyer Email", findFirstMatch(emailPattern, content));
        info.put("Supplier Email", findFirstMatch(supplierEmailPattern, content));

        return info;
    }

    private static String findFirstMatch(Pattern pattern, String content) {
        Matcher matcher = pattern.matcher(content);
        return matcher.find() ? matcher.group(1).trim() : "";
    }

    public static void exportWithHeaderAndItems(Map<String, String> headerInfo, List<LineItem> items, String excelPath) throws IOException {
        try (OutputStream os = new FileOutputStream(excelPath)) {
            ExcelWriter excelWriter = EasyExcel.write(os).build();
            WriteSheet sheet = EasyExcel.writerSheet("PO Lines").build();

            List<List<String>> headRows = new ArrayList<>();
            for (Map.Entry<String, String> entry : headerInfo.entrySet()) {
                headRows.add(Arrays.asList(entry.getKey(), entry.getValue()));
            }
            excelWriter.write(headRows, sheet);

            WriteTable table = EasyExcel.writerTable(1).head(LineItem.class).build();
            excelWriter.write(items, sheet, table);

            excelWriter.finish();
        }
    }
}
