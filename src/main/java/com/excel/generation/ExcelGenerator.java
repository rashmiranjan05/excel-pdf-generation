package com.excel.generation;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.DecimalFormat;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelGenerator {

    public static void main(String[] args) {
        try {
            // Create a new Workbook
            Workbook workbook = new XSSFWorkbook();

            // Create a new Sheet
            Sheet sheet = workbook.createSheet("Sheet1");
//            Font font = workbook.createFont();
//            font.setFontHeightInPoints((short)9);
//            font.setBold(true);

            // Create a header row
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Name", "Age", "Email"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Add data to the sheet
            String[][] data = {
                    {"John Doe", "30", "john@example.com"},
                    {"Jane Smith", "28", "jane@example.com"},
                    {"Bob Johnson", "35", "bob@example.com"}
            };
            int rowNum = 1;
            for (String[] rowData : data) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String value : rowData) {
                    row.createCell(colNum++).setCellValue(value);
                }
            }

            // Save the workbook to a file
            String filePath = "example.xlsx";
            FileOutputStream fileOut = new FileOutputStream(filePath);
            workbook.write(fileOut);
            fileOut.close();


            generatePdf();
            System.out.println("Excel file generated successfully.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public  static void generatePdf() throws DocumentException, FileNotFoundException {
        String pdfFilePath = "output.pdf";

        Document document;
        PdfPTable mainTable;
        PdfWriter writer;
        try {
            // Step 1: Create a Document object
            document = new Document();

            // Step 2: Create a PdfWriter to write the document to a file
            writer = PdfWriter.getInstance(document, new FileOutputStream(pdfFilePath));

            // Step 3: Open the document
            document.open();

            // Step 4: Add content to the document
//            String content = "Hello, this is a sample PDF generated using iText library in Java.";
//            document.add(new Paragraph(content));
            Font boldFont = FontFactory.getFont(FontFactory.HELVETICA, 8, Font.BOLD, new BaseColor(0, 0, 0));
            Font normalFont = FontFactory.getFont(FontFactory.HELVETICA, 7, Font.NORMAL, new BaseColor(0, 0, 0));
            Font headerFont = FontFactory.getFont(FontFactory.HELVETICA, 10, Font.BOLD, new BaseColor(0, 0, 0));
            DecimalFormat decimalFormatter = new DecimalFormat("$###,##0.00");

            PdfPTable headertable = new PdfPTable(2);
            headertable.setWidthPercentage(100);
            headertable.setSpacingBefore(10f);
            headertable.setSpacingAfter(10f);
            float[] widths = {2f, 5f};

                headertable.setWidths(widths);


            // PdfPCell imgCell = new PdfPCell(img, false);
//            imgCell.setHorizontalAlignment(Element.ALIGN_LEFT);
//            imgCell.setBorder(0);
//            headertable.addCell(imgCell);

            PdfPCell headerCell = new PdfPCell(
                    new Paragraph(
                            "Company Name: " + "businessName" + "\r\n\r\nFrom Date: " + headerFont));
            headerCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            headerCell.setBorder(0);
            headertable.addCell(headerCell);
            document.add(headertable);

            mainTable = new PdfPTable(10);
            mainTable.setWidthPercentage(100);
            mainTable.setSpacingBefore(10f);
            float[] paraWidths = {0.8f, 1.3f, 0.8f, 1f, 1.2f, 1.2f, 1f, 1.2f, 1.2f, 0.8f};
            mainTable.setWidths(paraWidths);

            PdfPCell dateHeadingCell = new PdfPCell(new Paragraph("Date", (com.itextpdf.text.Font) boldFont));
            dateHeadingCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            dateHeadingCell.setBorderWidthLeft(1f);
            dateHeadingCell.setBorderWidthRight(1f);
            dateHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(dateHeadingCell);

            PdfPCell vendorHeadingCell = new PdfPCell(new Paragraph("Vendor Name", (com.itextpdf.text.Font) boldFont));
            vendorHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            vendorHeadingCell.setBorderWidthLeft(0f);
            vendorHeadingCell.setBorderWidthRight(1f);
            vendorHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(vendorHeadingCell);

            PdfPCell poHeadingCell = new PdfPCell(new Paragraph("P.O", (com.itextpdf.text.Font) boldFont));
            poHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            poHeadingCell.setBorderWidthLeft(0f);
            poHeadingCell.setBorderWidthRight(1f);
            poHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(poHeadingCell);

            PdfPCell invNoHeadingCell = new PdfPCell(new Paragraph("Bill No.", (com.itextpdf.text.Font) boldFont));
            invNoHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            invNoHeadingCell.setBorderWidthLeft(0f);
            invNoHeadingCell.setBorderWidthRight(1f);
            invNoHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(invNoHeadingCell);

            PdfPCell paymentHeadingCell = new PdfPCell(new Paragraph("Payment Method", (com.itextpdf.text.Font) boldFont));
            paymentHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            paymentHeadingCell.setBorderWidthLeft(0f);
            paymentHeadingCell.setBorderWidthRight(1f);
            paymentHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(paymentHeadingCell);

            PdfPCell traceHeadingCell = new PdfPCell(new Paragraph("Check/Trace No.", (com.itextpdf.text.Font) boldFont));
            traceHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            traceHeadingCell.setBorderWidthLeft(0f);
            traceHeadingCell.setBorderWidthRight(1f);
            traceHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(traceHeadingCell);

            PdfPCell bankHeadingCell = new PdfPCell(new Paragraph("Funding Bank", (com.itextpdf.text.Font) boldFont));
            bankHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            bankHeadingCell.setBorderWidthLeft(0f);
            bankHeadingCell.setBorderWidthRight(1f);
            bankHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(bankHeadingCell);

            PdfPCell statusHeadingCell = new PdfPCell(new Paragraph("Status", (com.itextpdf.text.Font) boldFont));
            statusHeadingCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            statusHeadingCell.setBorderWidthLeft(0f);
            statusHeadingCell.setBorderWidthRight(1f);
            statusHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(statusHeadingCell);

            PdfPCell clearanceCell = new PdfPCell(new Paragraph("Payment Date", (com.itextpdf.text.Font) boldFont));
            clearanceCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            clearanceCell.setBorderWidthLeft(0f);
            clearanceCell.setBorderWidthRight(1f);
            clearanceCell.setPaddingBottom(2f);
            mainTable.addCell(clearanceCell);

            PdfPCell amtHeadingCell = new PdfPCell(new Paragraph("Amount", (com.itextpdf.text.Font) boldFont));
            amtHeadingCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            amtHeadingCell.setBorderWidthLeft(0f);
            amtHeadingCell.setBorderWidthRight(1f);
            amtHeadingCell.setPaddingBottom(2f);
            mainTable.addCell(amtHeadingCell);


            // Step 5: Close the document
            PdfPCell dateCell = new PdfPCell(new Paragraph("date", normalFont));
            dateCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            dateCell.setPaddingBottom(3f);
            dateCell.setBorderWidthLeft(1f);
            dateCell.setBorderWidthTop(0f);
            dateCell.setBorderWidthRight(1f);
            dateCell.setBorderWidthBottom(1f);
            mainTable.addCell(dateCell);

            PdfPCell vendorCell = new PdfPCell(new Paragraph("vendorName", normalFont));
            vendorCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            vendorCell.setPaddingBottom(3f);
            vendorCell.setBorderWidthLeft(0f);
            vendorCell.setBorderWidthTop(0f);
            vendorCell.setBorderWidthRight(1f);
            vendorCell.setBorderWidthBottom(1f);
            mainTable.addCell(vendorCell);

            PdfPCell poCell = new PdfPCell(new Paragraph("po", normalFont));
            poCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            poCell.setPaddingBottom(3f);
            poCell.setBorderWidthLeft(0f);
            poCell.setBorderWidthTop(0f);
            poCell.setBorderWidthRight(1f);
            poCell.setBorderWidthBottom(1f);
            mainTable.addCell(poCell);

            PdfPCell invNoCell = new PdfPCell(new Paragraph("invoiceNo", normalFont));
            invNoCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            invNoCell.setPaddingBottom(3f);
            invNoCell.setBorderWidthLeft(0f);
            invNoCell.setBorderWidthTop(0f);
            invNoCell.setBorderWidthRight(1f);
            invNoCell.setBorderWidthBottom(1f);
            mainTable.addCell(invNoCell);

            PdfPCell paymentMethodCell = new PdfPCell(new Paragraph("payment", normalFont));
            paymentMethodCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            paymentMethodCell.setPaddingBottom(3f);
            paymentMethodCell.setBorderWidthLeft(0f);
            paymentMethodCell.setBorderWidthTop(0f);
            paymentMethodCell.setBorderWidthRight(1f);
            paymentMethodCell.setBorderWidthBottom(1f);
            mainTable.addCell(paymentMethodCell);

            PdfPCell chkTraceCell = new PdfPCell(new Paragraph("chkTraceNo", normalFont));
            chkTraceCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            chkTraceCell.setPaddingBottom(3f);
            chkTraceCell.setBorderWidthLeft(0f);
            chkTraceCell.setBorderWidthTop(0f);
            chkTraceCell.setBorderWidthRight(1f);
            chkTraceCell.setBorderWidthBottom(1f);
            mainTable.addCell(chkTraceCell);

            PdfPCell bankNameCell = new PdfPCell(new Paragraph("bankName" + " (" + 12345 + ")", normalFont));
            bankNameCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            bankNameCell.setPaddingBottom(3f);
            bankNameCell.setBorderWidthLeft(0f);
            bankNameCell.setBorderWidthTop(0f);
            bankNameCell.setBorderWidthRight(1f);
            bankNameCell.setBorderWidthBottom(1f);
            mainTable.addCell(bankNameCell);

            PdfPCell statusCell = new PdfPCell(new Paragraph("status", normalFont));
            statusCell.setHorizontalAlignment(Element.ALIGN_LEFT);
            statusCell.setPaddingBottom(3f);
            statusCell.setBorderWidthLeft(0f);
            statusCell.setBorderWidthTop(0f);
            statusCell.setBorderWidthRight(1f);
            statusCell.setBorderWidthBottom(1f);
            mainTable.addCell(statusCell);

            PdfPCell cleranceCell = new PdfPCell(new Paragraph("processedDate", normalFont));
            cleranceCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            cleranceCell.setPaddingBottom(3f);
            cleranceCell.setBorderWidthLeft(0f);
            cleranceCell.setBorderWidthTop(0f);
            cleranceCell.setBorderWidthRight(1f);
            cleranceCell.setBorderWidthBottom(1f);
            mainTable.addCell(cleranceCell);

            PdfPCell amtCell = new PdfPCell(new Paragraph(decimalFormatter.format(12345), normalFont));
            amtCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
            amtCell.setPaddingBottom(3f);
            amtCell.setBorderWidthLeft(0f);
            amtCell.setBorderWidthTop(0f);
            amtCell.setBorderWidthRight(1f);
            amtCell.setBorderWidthBottom(1f);
            mainTable.addCell(amtCell);
        } catch (DocumentException e) {
            throw new RuntimeException(e);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }

        document.add(mainTable);
        document.close();
        writer.close();



        }
    }

