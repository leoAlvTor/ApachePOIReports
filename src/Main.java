import Controller.ControllerWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {
    public static void main(String[] args) {
        //CreateFile.createDocument("/home/leomnj/Documentos/archivo");
        String filePath = "/home/leomnj/Documentos/archivo.xlsx";

        ControllerWorkbook controller = new ControllerWorkbook();
        controller.readData("/home/leomnj/Documentos/archivo.xlsx");
        System.out.println("--------------------------------------------");
        XSSFWorkbook workbook = controller.readFile(filePath);
        XSSFSheet sheet1 = controller.readSheets(workbook)[1];

        Object[][] bookData = {
                {"Leo" , "Alvarado", "Torres", 10},
                {"Pedro" , "Illaisaca", "Jeje", 22},
                {"Xavier" , "Robles", "Ajio", 303},
                {"Israel" , "Chuchuca", "Hoho", 404},
        };

        if(controller.writeData(filePath, bookData, workbook, sheet1))
            System.out.println("DATA wrote correctly");
        else
            System.out.println("ERROR WHILE WRITING");
    }
}
