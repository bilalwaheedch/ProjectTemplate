package utility;

import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.bson.Document;

import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

/**
 * Created by Bilal on 07-02-2017.
 */
public class DataReader {
    /*
    Excel Reader Starts Here
     */

    HSSFWorkbook wb = null;
    HSSFSheet sheet = null;
    Cell cell = null;
    FileOutputStream fio = null;
    int numberOfRows, numberOfCol, rowNum;

    public String[][] fileReader(String path)throws IOException {
        String [] [] data = {};
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        wb = new HSSFWorkbook(fis);
        sheet = wb.getSheetAt(0);
        numberOfRows = sheet.getLastRowNum();
        numberOfCol =  sheet.getRow(0).getLastCellNum();
        data = new String[numberOfRows+1][numberOfCol+1];

        for(int i=1; i<data.length; i++){
            HSSFRow rows = sheet.getRow(i);
            for(int j=0; j<numberOfCol; j++){
                HSSFCell cell = rows.getCell(j);
                String cellData = excelGetCellValue(cell);
                data[i][j] = cellData;
            }
        }
        return  data;
    }
    public String[] excelColReader(String path, int col)throws IOException{
        String []  data = {};
        File file = new File(path);
        FileInputStream fis = new FileInputStream(file);
        wb = new HSSFWorkbook(fis);
        sheet = wb.getSheetAt(0);
        numberOfRows = sheet.getLastRowNum();
        numberOfCol =  col;
        data = new String[numberOfRows];

        for(int i=0; i<data.length; i++){
            HSSFRow rows = sheet.getRow(i+1);
            for(int j=0; j<numberOfCol; j++){
                HSSFCell cell = rows.getCell(j);
                String cellData = excelGetCellValue(cell);
                data[i] = cellData;
            }
        }
        return  data;
    }

    public String excelGetCellValue(HSSFCell cell){
        Object value = null;

        int dataType = cell.getCellType();
        switch(dataType){
            case HSSFCell.CELL_TYPE_NUMERIC:
                value = cell.getNumericCellValue();
                break;
            case HSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
        }
        return value.toString();

    }

    public void excelWriteBack(String value)throws IOException {
        wb = new HSSFWorkbook();
        sheet = wb.createSheet();
        Row row = sheet.createRow(rowNum);
        row.setHeightInPoints(10);

        fio = new FileOutputStream(new File("ExcelFile.xls"));
        wb.write(fio);
        for(int i=0; i<row.getLastCellNum(); i++){
            row.createCell(i);
            cell.setCellValue(value);
        }
        fio.close();
        wb.close();
    }
    /*
    Excel Reader Ends Here
     */

    /*
    DB Reader Starts Here
     */
    public static MongoDatabase mongoDatabase = null;
    public static MongoDatabase mongodb = null;


    Connection connect = null;
    Statement statement = null;
    PreparedStatement ps = null;
    ResultSet resultSet = null;

    public static Properties loadProperties() throws IOException{
        Properties prop = new Properties();
        InputStream ism = new FileInputStream("src/MySql.properties");
        prop.load(ism);
        ism.close();
        return prop;
    }

    public static Properties loadProperties(String name) throws IOException{
        Properties prop = new Properties();
        InputStream ism = new FileInputStream("src/config/"+name);
        prop.load(ism);
        ism.close();
        return prop;
    }

    public void connectToDatabase() throws IOException, SQLException, ClassNotFoundException {
        Properties prop = loadProperties("MySql.properties");
        String driverClass = prop.getProperty("MYSQLJDBC.driver");
        String url = prop.getProperty("MYSQLJDBC.url");
        String userName = prop.getProperty("MYSQLJDBC.userName");
        String password = prop.getProperty("MYSQLJDBC.password");
        Class.forName(driverClass);
        connect = DriverManager.getConnection(url,userName,password);
        //  System.out.println("Database is connected");

    }



    public static MongoDatabase connectMongoDB() throws IOException {
        try {
            Properties prop = loadProperties("mongodb.properties");
            String host = prop.getProperty("mongodn.url");
            MongoClientURI mongoClientURI = new MongoClientURI(host);
            MongoClient mongoClient = new MongoClient(mongoClientURI);
            System.out.println("MongoDB Connection Eastablished");
            mongoDatabase = mongoClient.getDatabase(prop.getProperty("mongodb.dbname"));
            System.out.println("Database Connected");
        }catch (IOException ex){
            System.out.println(ex.getMessage());
        }
        return mongoDatabase;
    }


    public List<String> sqlReadDataBase(String tableName, String columnName)throws Exception{
        List<String> data = new ArrayList<String>();

        try {
            connectToDatabase();
            statement = connect.createStatement();
            resultSet = statement.executeQuery("select * from " + tableName);
            data = getResultSetData(resultSet, columnName);
        } catch (ClassNotFoundException e) {
            throw e;
        }finally{
            close();
        }
        return data;
    }


    private void close() {
        try{
            if(resultSet != null){
                resultSet.close();
            }
            if(statement != null){
                statement.close();
            }
            if(connect != null){
                connect.close();
            }
        }catch(Exception e){

        }
    }


    private List<String> getResultSetData(ResultSet resultSet2, String columnName) throws SQLException {
        List<String> dataList = new ArrayList<String>();
        while(resultSet.next()){
            String itemName = resultSet.getString(columnName);
            dataList.add(itemName);
        }
        return dataList;
    }

    // function  for Data insert into MySQL Database
    public void InsertDataFromArryToMySql(int [] ArrayData,String tableName, String columnName)
    //InsertDataFromArryListToMySql

    //  public void InsertDataFromArryToMySql()
    {

        try {
            connectToDatabase();

            //  connect.createStatement("INSERT into tbl_insertionSort set SortingNumbers=1000");

            for(int n=0; n<ArrayData.length; n++){

                ps = connect.prepareStatement("INSERT INTO "+tableName+" ( "+columnName+" ) VALUES(?)");
                ps.setInt(1,ArrayData[n]);
                ps.executeUpdate();
                //System.out.println(list[n]);
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
        //connection = ConnectionConfiguration.getConnection();
    }


    // Function for Insert Single value in a table

    public void sqlInsertDataFromStringToMySql(String ArrayData,String tableName, String columnName)


    //  public void InsertDataFromArryToMySql()
    {

        try {
            connectToDatabase();

            //  connect.createStatement("INSERT into tbl_insertionSort set SortingNumbers=1000");


            ps = connect.prepareStatement("INSERT INTO "+tableName+" ( "+columnName+" ) VALUES(?)");
            ps.setString(1,ArrayData);
            ps.executeUpdate();
            //System.out.println(list[n]);


        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
        //connection = ConnectionConfiguration.getConnection();
    }




    public List<String> sqlDirectDatabaseQueryExecute(String passQuery,String dataColumn)throws Exception{
        List<String> data = new ArrayList<String>();

        try {
            connectToDatabase();
            statement = connect.createStatement();
            resultSet = statement.executeQuery(passQuery);
            data = getResultSetData(resultSet, dataColumn);
        } catch (ClassNotFoundException e) {
            throw e;
        }finally{
            close();
        }
        return data;
    }

//

    public void sqlInsertDataFromArryListToMySql(List<Object> list,String tableName, String columnName)
    //InsertDataFromArryListToMySql

    //  public void InsertDataFromArryToMySql()
    {

        try {
            connectToDatabase();

            //  connect.createStatement("INSERT into tbl_insertionSort set SortingNumbers=1000");

            for(Object st:list){
                // System.out.println(st);

                ps = connect.prepareStatement("INSERT INTO "+tableName+" ( "+columnName+" ) VALUES(?)");
                ps.setObject(1,st);
                ps.executeUpdate();
                //System.out.println(list[n]);
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
        //connection = ConnectionConfiguration.getConnection();
    }


    public ResultSet Query(String sql)  throws SQLException {
        Connection conn = null;
        Statement stmt = null;
        ResultSet resultSet = null;
        try {
            connectToDatabase();
            statement = connect.createStatement();
            resultSet = statement.executeQuery(sql);
        } catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        }

        return resultSet;
    }


    public static List<Document> mongoGetMongoDBDataDocument(String collectionName ) {
        List<Document> documents = null;
        MongoCollection<Document> collection = null;
        try {
            collection = mongoDatabase.getCollection(collectionName);
            documents = (List<Document>) collection.find().into(
                    new ArrayList<Document>());

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return documents;
    }
    /*
    DB Reader Ends Here
     */
}
