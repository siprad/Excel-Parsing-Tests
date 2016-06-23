package com.siprad;

import com.incesoft.tools.excel.xlsx.*;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Scanner;


public class Main {

    public static String filename = "e.xlsx";

    public static void addStyleAndRichText(SimpleXLSXWorkbook wb, Sheet sheet)
            throws Exception {
        Font font2 = wb.createFont();
        font2.setColor("FFFF0000");
        Fill fill = wb.createFill();
        fill.setFgColor("FF00FF00");
        CellStyle style = wb.createStyle(font2, fill);

        RichText richText = wb.createRichText();
        richText.setText("test_text");
        Font font = wb.createFont();
        font.setColor("FFFF0000");
        richText.applyFont(font, 1, 2);
        sheet.modify(0, 0, (String) null, style);
        sheet.modify(1, 0, richText, null);
    }

    static public void addRecordsOnTheFly(SimpleXLSXWorkbook wb, Sheet sheet,
                                          int rowOffset) {
        int columnCount = 10;
        int rowCount = 10;
        int offset = rowOffset;
        for (int r = offset; r < offset + rowCount; r++) {
            int modfiedRowLength = sheet.getModfiedRowLength();
            for (int c = 0; c < columnCount; c++) {
                sheet.modify(modfiedRowLength, c, r + "," + c, null);
            }
        }
    }

    private static void printRow(int rowPos, Cell[] row) {
        int cellPos = 0;
        try {
            for (Cell cell : row) {
                System.out.println(Sheet.getCellId(rowPos, cellPos) + "="
                        + cell.getValue());
                cellPos++;
            }
        }
        catch (Exception e){

        }
    }

    public static void testLoadALL(SimpleXLSXWorkbook workbook) {
        // medium data set,just load all at a time
        Sheet sheetToRead = workbook.getSheet(0);
        List<Cell[]> rows = sheetToRead.getRows();
        int rowPos = 0;
        for (Cell[] row : rows) {
            printRow(rowPos, row);
            rowPos++;
        }
    }

    public static void testIterateALL(SimpleXLSXWorkbook workbook) {
        // here we assume that the sheet contains too many rows which will leads
        // to memory overflow;
        // So we get sheet without loading all records
        Sheet sheetToRead = workbook.getSheet(0, false);
        Sheet.SheetRowReader reader = sheetToRead.newReader();
        Cell[] row;
        int rowPos = 0;
        while ((row = reader.readRow()) != null) {
            printRow(rowPos, row);
            rowPos++;
        }
    }

    public static void testWrite(SimpleXLSXWorkbook workbook,
                                 OutputStream outputStream) throws Exception {
        Sheet sheet = workbook.getSheet(0);
        addRecordsOnTheFly(workbook, sheet, 0);
        workbook.commit(outputStream);
    }

    /**
     * Commit serveral times for large data set
     *
     * @param workbook
     * @param output
     * @throws Exception
     */
    public static void testWriteByIncrement(SimpleXLSXWorkbook workbook,
                                            OutputStream output) throws Exception {
        SimpleXLSXWorkbook.Commiter commiter = workbook.newCommiter(output);
        commiter.beginCommit();

        Sheet sheet = workbook.getSheet(0, false);
        commiter.beginCommitSheet(sheet);
        addRecordsOnTheFly(workbook, sheet, 0);
        commiter.commitSheetWrites();
        addRecordsOnTheFly(workbook, sheet, 20);
        commiter.commitSheetWrites();
        addRecordsOnTheFly(workbook, sheet, 40);
        commiter.commitSheetWrites();
        commiter.endCommitSheet();

        commiter.endCommit();
    }

    /**
     * first, modify the original sheet; and then append some data
     *
     * @param workbook
     * @param output
     * @throws Exception
     */
    public static void testMergeBeforeWrite(SimpleXLSXWorkbook workbook,
                                            OutputStream output) throws Exception {
        Sheet sheet = workbook.getSheet(0, false);// assuming original data
        // set is large
        addStyleAndRichText(workbook, sheet);
        addRecordsOnTheFly(workbook, sheet, 5);

        SimpleXLSXWorkbook.Commiter commiter = workbook.newCommiter(output);
        commiter.beginCommit();
        commiter.beginCommitSheet(sheet);
        // merge it first,otherwise the modification will not take effect
        commiter.commitSheetModifications();

        // row = -1, for appending after the last row
        sheet.modify(-1, 1, "append1", null);
        sheet.modify(-1, 2, "append2", null);
        // lets assume there are many rows here...
        commiter.commitSheetWrites();// flush writes,save memory

        sheet.modify(-1, 1, "append3", null);
        sheet.modify(-1, 2, "append4", null);
        // lets assume there are many rows here,too ...
        commiter.commitSheetWrites();// flush writes,save memory

        commiter.endCommitSheet();
        commiter.endCommit();
    }

    private static SimpleXLSXWorkbook newWorkbook() {

        return new SimpleXLSXWorkbook(new File(filename));
    }

    private static OutputStream newOutput(String suffix) throws Exception {
        return new BufferedOutputStream(new FileOutputStream("/sample_"
                + suffix + ".xlsx"));
    }

    public static void main(String[] args) throws Exception {

       if(args.length>0)
           filename= args[0]+".xlsx";
        Scanner reader = new Scanner(System.in);

        System.out.println("sjxlsx- enter y to start");
        reader.next();

        Long loadStartTime = System.currentTimeMillis();
        System.out.println("STARTING EXCEL PARSE");
        try {
            SimpleXLSXWorkbook workbook = newWorkbook();
            Long loadEndTime = System.currentTimeMillis();

            System.out.println("Starting Writing");
        // READ by classic mdoe - load all records
          //  testLoadALL(workbook);
        // READ by stream mode - iterate records one by one
            testIterateALL(workbook);
            Long writeEndTime = System.currentTimeMillis();


            System.out.println("sjxlsx-enter y to stop");
            reader.next();
            System.out.println("Load time:\t"+(loadEndTime-loadStartTime));
            System.out.println("Write time:\t"+(writeEndTime-loadEndTime));
            System.out.println("Total time:\t"+(writeEndTime-loadStartTime));

        }
        catch (Exception e){
            e.printStackTrace();
            System.out.println("ERR");
        }
     /*   // WRITE - we take WRITE as a special kind of MODIFY
       OutputStream output = newOutput("write");
        //testWrite(workbook, output);
        //output.close();

        // WRITE large data
        output = newOutput("write_inc");
        testWriteByIncrement(workbook, output);
        output.close();

        // MODIFY it and WRITE large data
        output = newOutput("merge_write");
        testMergeBeforeWrite(workbook, output);
        output.close();*/
    }
}
