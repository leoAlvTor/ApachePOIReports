package Controller;

import java.sql.*;

public class Database{

    private static final String DRIVER = "com.mysql.cj.jdbc.Driver";
    private static final String URL = "jdbc:mysql://localhost:3306/mundoganaderomercy";
    private static final String USER = "monty";
    private static final String PASS = "montypassword";
    private static Database jdbc1 = null;
    private Statement statement = null;
    private static Connection connection = null;

    public Database(){
        connect();
    }

    public void connect(){
        try{
            Class.forName(DRIVER);
            connection = DriverManager.getConnection(URL, USER, PASS);
            this.statement = connection.createStatement();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public ResultSet query(String sql){
        try{
            return this.statement.executeQuery(sql);
        }catch (SQLException e){
            System.out.println(e.getMessage() + "<-- ERROR");
        }
        return null;
    }

    public static Database getDatabase(){
        if(jdbc1 == null){
            jdbc1 = new Database();
        }
        return jdbc1;
    }

    public static void disconnect(){
        try{
            if(!connection.isClosed()) {
                connection.close();
            }
        }catch(Exception e){
            e.printStackTrace();
        }
    }

}
