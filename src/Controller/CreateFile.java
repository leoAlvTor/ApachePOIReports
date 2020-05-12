package Controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class CreateFile {
    public CreateFile(){}

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
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(documentName+".xlsx");
        }catch (FileNotFoundException fileNotFoundException){
            System.out.println("Error while creating document!");
            return false;
        }
        // Creating document sheets
        Sheet sheet = workbook.createSheet("Sheet1");
        Sheet graficas = workbook.createSheet("Grafica");
        try {
            workbook.write(outputStream);
            System.out.println("Document sheet created!");
            return true;
        }catch (IOException ioException){
            System.out.println("Error while writing on document!");
            return false;
        }
    }
}
