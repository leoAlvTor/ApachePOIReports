import Controller.ControllerWorkbook;
import Controller.Database;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;


public class Main {
    public static void main(String[] args) {
        Object[] titles = new Object[]{"Nombre Del Proveedor", "Nombre Del Producto", "Cantidad Vendida"};
        nddend(testing(), titles);
    }

    private static void nddend(List<List<Object>> objectsLists, Object[] titles){
        //CreateFile.createDocument("/home/leomnj/Documentos/archivo");
        String filePath = "/home/leomnj/Documentos/archivo.xlsx";

        ControllerWorkbook controller = new ControllerWorkbook();
        controller.readData("/home/leomnj/Documentos/archivo.xlsx");
        System.out.println("--------------------------------------------");
        XSSFWorkbook workbook = controller.readFile(filePath);
        controller.writeColumnTitleWithStyle(titles, workbook, 0);
        XSSFSheet sheet1 = controller.readSheets(workbook)[0];

        if(controller.writeData(filePath, objectsLists, workbook, sheet1))
            System.out.println("DATA wrote correctly");
        else
            System.out.println("ERROR WHILE WRITING");
    }

    public static List<List<Object>> testing(){
        Database database = Database.getDatabase();

        String productosMasVendidos = "select proveedores.prov_nombre, productos.prod_name, sum(factura_detalle.cantidad) " +
                "from productos, factura_detalle, factura_cabecera, clientes, proveedores " +
                "where factura_detalle.prod_id = productos.prod_id " +
                "and factura_cabecera.factura_id = factura_detalle.cabecera_id " +
                "and clientes.cliente_id = factura_cabecera.cliente_ruc " +
                "and proveedores.prov_ruc = productos.prov_RUC " +
                "group by 1 " +
                "order by 3 desc";

        ResultSet rs = database.query(productosMasVendidos);

        try{
            List<List<Object>> objectLists = new ArrayList<>();
            List<Object> objectList;
            while(rs.next()){
                objectList = new ArrayList<>();
                objectList.add(rs.getString(1));
                objectList.add(rs.getString(2));
                objectList.add(rs.getDouble(3));
                objectLists.add(objectList);
            }
            return objectLists;
        }catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }
}
