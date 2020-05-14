package Controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ControllerWorkbook {
    public ControllerWorkbook(){}

    /**
     * Create a new Excel document with one sheet on it.
     *
     * @param documentName The parameter is the name of the document.
     * @return The method return True if document and sheet were created or false if otherwise.
     */
    public static boolean createDocument(String documentName) {
        // Creating a Workbook instance.
        Workbook workbook = new HSSFWorkbook();
        // Creating an OutputStream (File Creation)
        OutputStream outputStream ;
        try {
            // Create a new workbook.
            outputStream = new FileOutputStream(documentName+".xlsx");
        }catch (FileNotFoundException fileNotFoundException){
            System.out.println("Error while creating document!");
            return false;
        }
        // Creating document sheets
        Sheet sheet = workbook.createSheet("Sheet1");
        Sheet graficas = workbook.createSheet("Grafica");
        try {
            // Writing sheets to the workbook.
            workbook.write(outputStream);
            System.out.println("Document sheet created!");
            return true;
        }catch (IOException ioException){
            System.out.println("Error while writing on document!");
            return false;
        }
    }

    public XSSFWorkbook readFile(String filePath){
        // Creating an instance of the file by it's name.
        File file = new File(filePath);
        // Reading stream of bytes
        FileInputStream fileInputStream;
        try {
            fileInputStream = new FileInputStream(file);
            return new XSSFWorkbook(fileInputStream);
        }catch (FileNotFoundException fileNotFoundException){
            System.out.println("The file could not be found!");
            return null;
        }catch(IOException ioException){
            System.out.println("Error while reading the file!");
            return null;
        }
    }

    public XSSFSheet[] readSheets(XSSFWorkbook workbook){
        XSSFSheet[] sheets = new XSSFSheet[workbook.getNumberOfSheets()];
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            sheets[i] = workbook.getSheetAt(i);
        }
        return sheets;
    }

    public Iterator<Row> readRowsFromSheet(XSSFSheet xssfSheet){
        return xssfSheet.iterator();
    }

    public List<Cell> readCells(Iterator<Row> iterator){
        List<Cell> cellList = new ArrayList<>();
        while(iterator.hasNext()){
            Row row =iterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                cellList.add(cell);
            }
        }
        return cellList;
    }

    public void printCellData(List<Cell> cells){
        for (Cell cell: cells){
            switch (cell.getCellType()){
                case STRING -> System.out.println(cell.getStringCellValue());
                case NUMERIC -> System.out.println(cell.getNumericCellValue());
                case BOOLEAN -> System.out.println(cell.getBooleanCellValue());
            }
        }
    }

    public void readData(String path){
        // Reading workbook.
        XSSFWorkbook workbook = readFile(path);
        System.out.println("Workbook read correctly!");

        // Sheet reading.
        XSSFSheet[] sheets = readSheets(workbook);
        System.out.println("Sheets read correctly");

        // Iterating over sheets.
        for (XSSFSheet sheet : sheets) {
            Iterator<Row> rows = readRowsFromSheet(sheet);
            List<Cell> cells = readCells(rows);
            printCellData(cells);
        }

    }

    public boolean writeData(String filePath, Object[][] dataObjects, XSSFWorkbook workbook, XSSFSheet sheet){
        int rowCount = 0;
        for(Object[]  data: dataObjects){
            Row row = sheet.createRow(++rowCount);
            int columnCount = 0;

            for(Object field : data){
                Cell cell = row.createCell(++columnCount);
                if(field instanceof String)
                    cell.setCellValue((String) field);
                else if(field instanceof Integer)
                    cell.setCellValue((Integer) field);
                else if(field instanceof Double)
                    cell.setCellValue((Double) field);
            }
        }
        try{
            FileOutputStream outputStream = new FileOutputStream(filePath);
            workbook.write(outputStream);
            return true;
        }catch (FileNotFoundException fileNotFoundException){
            System.out.println("File not found in the specified path: " + filePath);
            return false;
        }catch (IOException ioException){
            System.out.println("Error while writing file!");
            return false;
        }
    }

}
