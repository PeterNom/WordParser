package org.example;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.

import java.io.*;
import java.util.List;
import java.util.Scanner;

// Importing Apache POI package
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) throws Exception {
        String[] rank = new String[]{"ΤΧΗΣ", "ΛΓΟΣ", "ΥΠΛΓΟΣ", "ΑΝΘΛΓΟΣ", "ΔΕΑ", "ΑΝΘΣΤΗΣ", "ΑΛΧΙΑΣ", "ΕΠΧΙΑΣ", "ΛΧΙΑΣ", "OBA", "ΣΤΡ"};
        int rowCount = 0;
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Personal");


        // Step 1: Getting path of the current working
        // directory where the Word document is located
        String path = System.getProperty("user.dir");
        path = path + File.separator + "test.docx";

        // Step 2: Creating a file object with the above
        // specified path.
        FileInputStream fin = new FileInputStream(path);

        // Step 3: Creating a document object for the Word
        // document.
        XWPFDocument document = new XWPFDocument(fin);

        // Step 4: Using the getParagraphs() method to
        // retrieve the list of paragraphs from the Word
        // file.
        List<XWPFTable> tables = document.getTables();

        // Step 5: Iterating through the list of paragraphs
        for (XWPFTable tableRow : tables) {
            String table = tableRow.getText();

            Scanner sc = new Scanner(table);
            String line;
            try {
                 do
                 {
                     line = sc.nextLine();
                     for (String s : rank) {
                        if (line.contains(s)) {
                            try (Scanner scanner = new Scanner(line)) {
                                String temp;

                                while (scanner.hasNext()) {
                                    temp = scanner.next();
                                    if (temp.equals(s)) {
                                        int colcnt = 0;
                                        XSSFRow row = sheet.createRow(rowCount++);
                                        XSSFCell cell = row.createCell(colcnt++);

                                        String temp2 = scanner.next();
                                        String temp3 = scanner.next();



                                        cell.setCellValue(temp);
                                        cell = row.createCell(colcnt++);
                                        cell.setCellValue(temp2);
                                        cell = row.createCell(colcnt++);
                                        cell.setCellValue(temp3);

                                    }
                                }
                            }
                        }
                    }
                } while (!line.isEmpty());
            }
            catch (Exception e)
            {
                System.out.println("Exception thrown: " + e);
            }
            sc.close();
        }
        try (FileOutputStream outputStream = new FileOutputStream("JavaBooks1.xlsx")) {
            workbook.write(outputStream);
        }
        finally {
            workbook.close();
            document.close();
        }
    }
    }