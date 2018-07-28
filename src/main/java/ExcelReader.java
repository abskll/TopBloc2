import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;



public class ExcelReader {
    //public static final String SAMPLE_XLS_FILE_PATH = "./sample-xls-file.xls";
    public static final String WORKBOOK_1_PATH = "./Data1.xlsx";
    public static final String WORKBOOK_2_PATH = "./Data2.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbooknum1 = WorkbookFactory.create(new File(WORKBOOK_1_PATH));
        Workbook workbooknum2 = WorkbookFactory.create(new File(WORKBOOK_2_PATH));
        
        //String [][] life2 = new String[29][120];
        String [][] workbook1 = new String[5][3];
        String [][] workbook2 = new String[5][3];
        Integer [] numberSetOne = new Integer[4];
        Double [] numberSetTwo = new Double[4];
        String [] wordSetOne = new String[4];
        int numrow = 0;
        int col = 0;
        
       
        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbooknum1.getSheetAt(0);
        Sheet sheet2 = workbooknum2.getSheetAt(0);
        
        
        
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        
        /*
         * 
        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }


		*/
        
        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
                workbook1[numrow][col] = cellValue;
                col++;
            }
            col = 0;
            numrow++;
            System.out.println();
        }
        System.out.print("cols:" + col);
        System.out.print("numrows: " + numrow);
        numrow=0;
        col = 0;
        
        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet2) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
                workbook2[numrow][col] = cellValue;
                col++;
            }
            col = 0;
            numrow++;
            System.out.println();
        }
        
        
        //String ele;
        for (String rowele[]: workbook1) {
        	for (String ele: rowele) {
        		System.out.print(ele + "\t");
        	}
        	System.out.println();
        }
        
        
        //System.out.println(workbook2[4][0]);
        //System.out.println(workbook2[4][1]);//12 start column, 1 start row
        //System.out.println(workbook2[4][2]);
        
        
        
        for(int numele = 1; numele < workbook2.length - 1; numele++) {
        	numberSetOne[numele] = Integer.parseInt(workbook1[numele][0]) * Integer.parseInt(workbook2[numele][0]);
        }
        
        //System.out.println(numberSetOne[1]);
        for(int numele = 1; numele < workbook2.length - 1; numele++) {
        	numberSetTwo[numele] = (double) (Integer.parseInt(workbook1[numele][1]) / Integer.parseInt(workbook2[numele][1]));
        }
        for(int numele = 1; numele < workbook2.length - 1; numele++) {
        	wordSetOne[numele] = workbook1[numele][2] + workbook2[numele][2];
        }
        
       
        System.out.print(numberSetOne[2]);
        System.out.print(numberSetTwo[2]);
        System.out.print(wordSetOne[2]);
    }

  
}
