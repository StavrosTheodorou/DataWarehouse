package data;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class DB_Manager 
{
    private Connection con;
    
    public DB_Manager() 
    { 
        con = null;
        System.out.println("test");
    }
    
    private void connect()
    {
        try
        {
            Class.forName("org.sqlite.JDBC");
            
            Connection con_attempt = DriverManager.getConnection("jdbc:sqlite:4901le.db");
            
            System.out.println("Successful connection to DB"); 
            
            con = con_attempt;
        }
        catch(Exception e)
        {
            System.out.println("Error during connection to DB");
            e.printStackTrace();
            
            con = null;
        }
    }
    
    private void disconnect()
    {
        try 
        { 
            if (con != null) 
            {
                con.close();
                
                System.out.println("Successful disconnection from DB\n");
            }
            else
            {
                
            }
        } 
        catch(Exception e) 
        { 
            System.out.println("Error during disconnection from DB"); 
            e.printStackTrace(); 
        }
    }
    
    public void Insert(String table, String values)
    {
        connect();
        
        Statement statement = null;
        
        String query =  "INSERT INTO " + table + " (aa, arithMerid, arithOnom, perigrafiIdwn, monMetr, posotita, io, paratiriseis, barcode) VALUES(" + values + ");";
        
        try 
        {
            statement = con.createStatement();
            con.setAutoCommit(false);
            
            statement.executeUpdate(query);
            con.commit();
            
            System.out.println("Query execution: " + query + " (SUCCESS)"); 
            
            statement.close();
        }
        catch(SQLException e)
        {
            System.out.println("Query execution: " + query + " (FAILURE)");
            e.printStackTrace();
        }
        finally { disconnect(); }
    }
    
    public ArrayList<Object[]> Select(String columns, String table, String condition)
    {
        connect();
        
        ArrayList<Object[]> selectList = new ArrayList<Object[]>();
        
        Statement statement = null;
        
        String query =  "SELECT " + columns + " FROM " + table;
        
        query += condition.equals("") ? ";" : " WHERE " + condition + ";";
        
        try 
        {
            statement = con.createStatement();
            
            ResultSet set = statement.executeQuery(query);
            
            ResultSetMetaData rsmd = set.getMetaData();
                
            int columnsNumber = rsmd.getColumnCount();
            
            while (set.next()) 
            {
                Object[] row = new Object[columnsNumber];
                
                for (int i = 1; i <= columnsNumber; i++) row[i - 1] = set.getString(i);
                
                selectList.add(row);
            }
            
            System.out.println("Query execution: " + query + " (SUCCESS)");
            
            statement.close();
            set.close();
        } 
        catch(SQLException e ) 
        { 
            System.out.println("Query execution: " + query + " (FAILURE)");
            e.printStackTrace();
        } 
        finally { disconnect(); }
        
        return selectList;
    }
    
    public void Update(String table, String column, String new_value, String condition)
    {
        connect();
        
        Statement statement = null;
        
        String query =  "UPDATE " + table + " SET " + column + " = " + new_value + " WHERE " + condition + ";";
        
        try 
        {
            statement = con.createStatement();
            con.setAutoCommit(false);
            
            statement.executeUpdate(query);
            con.commit();
            
            System.out.println("Query execution: " + query + " (SUCCESS)"); 
            
            statement.close();
        }
        catch(SQLException e)
        {
            System.out.println("Query execution: " + query + " (FAILURE)");
            e.printStackTrace();
        }
        finally { disconnect(); }
    }
    
    public void Delete(String table, String condition)
    {
        connect();
        
        Statement statement = null;
        
        String query =  "DELETE FROM " + table + " WHERE " + condition + ";";
        
        try 
        {
            statement = con.createStatement();
            
            statement.executeUpdate(query);
            
            System.out.println("Query execution: " + query + " (SUCCESS)"); 
            
            statement.close();
        }
        catch(SQLException e)
        {
            System.out.println("Query execution: " + query + " (FAILURE)");
            e.printStackTrace();
        }
        finally { disconnect(); }
    }
    
    public Object Function(String function, String table, String columns, String condition)
    {
        connect();
        
        Object result = null;
        
        Statement statement = null;
        
        String query =  "SELECT " + function + "(" + columns + ") AS result FROM " + table;
        
        query += condition.equals("") ? ";" : " WHERE " + condition + ";";
        
        try 
        {
            statement = con.createStatement();
            
            ResultSet set = statement.executeQuery(query);
                
            set.next();
            
            result =  set.getObject("result");
            
            System.out.println("Query execution: " + query + " (SUCCESS)");
            
            statement.close();
            set.close();
        } 
        catch(SQLException e ) 
        { 
            System.out.println("Query execution: " + query + " (FAILURE)");
            e.printStackTrace();
        } 
        finally { disconnect(); }
        
        return result;
    }
    
}
