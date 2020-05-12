import Controller.CreateFile;

public class Main {
    public static void main(String[] args) {
        //CreateFile.createDocument("/home/leomnj/Documentos/archivo");

        CreateFile createFile = new CreateFile();
        createFile.readData("/home/leomnj/Documentos/archivo.xlsx");
    }
}
