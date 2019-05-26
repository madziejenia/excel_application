package excel.application;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class App
{
    private static String[] columns = {"Name of course", "Day of Week", "Local Time", "Number of hours", "Start Date"};
    private static List<Lessons> lessons = new ArrayList<>();



    static {
        lessons.add(new Lessons("Java basics", "Monday", "18:30 - 20:30", "2", "10/06/2019"));
        lessons.add(new Lessons("Java - Objects, Classes", "Thursday", "18:30 - 20:30", "2", "13/06/2019"));
        lessons.add(new Lessons("Java - Methods, Constructors", "Friday", "18:30 - 20:30", "2", "14/06/2019"));
        lessons.add(new Lessons("Inheritance", "Monday", "18:30 - 20:30", "2", "17/06/2019"));
        lessons.add(new Lessons("Polymorphism", "Wednesday", "18:30 - 20:30", "2", "19/06/2019"));
        lessons.add(new Lessons("SQL essentials", "Friday", "18:30 - 20:30", "2", "21/06/2019"));

    }

    public static void main(String[] args) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Lessons schedule");
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 16);
        headerFont.setColor(IndexedColors.BLACK.getIndex());


        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columns.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(headerCellStyle);
            cell.setCellValue(columns[i]);


        }



        int rowNumber = 1;

        for (Lessons lessons : lessons){
            Row row = sheet.createRow(rowNumber++);
            row.createCell(0).setCellValue(lessons.NameOfCourse);
            row.createCell(1).setCellValue(lessons.DayOfWeek);
            row.createCell(2).setCellValue(lessons.LocalTime);
            row.createCell(3).setCellValue(lessons.NumberOfHours);
            row.createCell(4).setCellValue(lessons.StartDate);

        }

        for (int i = 0; i < columns.length; i++){
            sheet.autoSizeColumn(i);
        }
            int rowCount = 0;
            Row rowTotal = sheet.createRow(rowCount + 7);
            Cell cellTotalText = rowTotal.createCell(0);
            cellTotalText.setCellValue("Total hours:");

            Cell cellTotal = rowTotal.createCell(3);
            cellTotal.setCellFormula("SUM(D2:D7)");


            FileOutputStream fileOut = new FileOutputStream("lessons.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

    }


}

