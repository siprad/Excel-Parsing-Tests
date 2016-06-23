package com.siprad;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Scanner;

public class Main {
    public static String filename = "e.xlsx";
    static XSSFRow row;
    public static void main(String[] args) throws Exception {

        if(args.length>0)
            filename= args[0]+".xlsx";

        Scanner reader = new Scanner(System.in);

        FileInputStream file = new FileInputStream(new File(filename));
        System.out.println("POI- enter y to start");
        reader.next();

        Long loadStartTime = System.currentTimeMillis();
        System.out.println("STARTING EXCEL PARSE");

        XSSFWorkbook workbook = new XSSFWorkbook(file);

       // OPCPackage pkg = OPCPackage.openOrCreate(new File(filename));
        //XSSFWorkbook workbook = new XSSFWorkbook(pkg);
                       //new XSSFWorkbook(new File(filename));
        Long loadEndTime = System.currentTimeMillis();

        System.out.println("Starting Writing");
        XSSFSheet spreadsheet =workbook.getSheetAt(0);

        Iterator<Row> rowIterator = spreadsheet.iterator();
        while (rowIterator.hasNext())
        {
            row = (XSSFRow) rowIterator.next();
            Iterator <Cell> cellIterator = row.cellIterator();
            while ( cellIterator.hasNext())
            {
                Cell cell = cellIterator.next();
                switch (cell.getCellType())
                {
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(
                                cell.getNumericCellValue() + " \t\t " );
                        break;
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(
                                cell.getStringCellValue() + " \t\t " );
                        break;
                }
            }
            System.out.println();
        }
        Long writeEndTime = System.currentTimeMillis();

        System.out.println();

        System.out.println("POI-enter y to stop");
        reader.next();
        System.out.println("Load time:\t"+(loadEndTime-loadStartTime));
        System.out.println("Write time:\t"+(writeEndTime-loadEndTime));
        System.out.println("Total time:\t"+(writeEndTime-loadStartTime));

    }
}
