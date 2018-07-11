/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this kjkkm, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication18;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SIKU
 */
public class SqlInputer {
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	Date date = new Date();
    Connection con;PreparedStatement st;
    int row = 1;
    double Copper = 0;  double GSW =0;  double PVC = 0; double Aluminum = 0;
      
    String cop;String url = "jdbc:ucanaccess://Reporting1.accdb";
    public void Put(String ItemCode, String Stage, String Name, String len, double Cu, double PVC, double GSW, String Machine, double Aluminum) throws ClassNotFoundException {
        try {

            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

            String SqlQuery = "INSERT INTO Output1([ID],[ItemCode],[Description],[Stage],[Kilometers],[Cu],[PVC],[GSW],[Machine],[Aluminum]) "
                    + " VALUES(?,?,?,?,?,?,?,?,?,?)";

            con = DriverManager.getConnection(url);
            st = con.prepareStatement(SqlQuery);
            st.setString(1, Machine + Name + date.toString());
            st.setString(2, ItemCode);
            st.setString(3, Name);
            st.setString(4, Stage);
            st.setString(5, len);
            st.setString(6, String.format("%.2f", Cu));
            st.setString(7, String.format("%.2f", PVC));
            st.setString(8, String.format("%.2f", GSW));
            st.setString(9, Machine);
            st.setString(10,String.format("%.2f", Aluminum ));

            st.executeUpdate();
            st.close();

        } catch (SQLException wx) {
            System.out.println(wx.getLocalizedMessage());
        }
    }
        Statement si;
        ResultSet rs;
    public void read() {
 String[] Machiness = {"HRB","HINTER","H8WIRE","TML","NAMPTON","37STRANDER","GODDERIDGE","BUNCHER",	   
"SPIRKA BRAIDER","F 13","18STRANDER","24STRANDER","61STRANDER","PS COILER","FS CORELINE","BM80 No1","BM80 No2",	   
"NM100","SHEATHLINE","DS130","Quadder","Twinner","PG1000","UNIT STRANDER","DTLAY UP","DT ARM","42s",   
"GG REWINDER", "ANDUART"};
        

        try {
            File file = new File("SCrap101.xlsx");
            FileInputStream input = new FileInputStream(file);
            XSSFWorkbook wb = (XSSFWorkbook) WorkbookFactory.create(input);
            XSSFSheet report = wb.getSheetAt(0);
            XSSFSheet actualreport = wb.getSheetAt(1);
            int m = 38;   int ss = 5;

            for (int i = 0; i < Machiness.length; i++) {
                 double TotalCopper = 0;double TotalPVC = 0;double TotalGSW = 0;double TotalAluminum = 0;
                
                String sqlquery = " SELECT * FROM Output1 WHERE Machine ='" + Machiness[i] + "'";

                Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
                con = DriverManager.getConnection(url);
                si = con.createStatement();
                rs = si.executeQuery(sqlquery);
               
                while (rs.next()) {
                    XSSFRow row1 = report.createRow(m);
                    Cell Cutxt = row1.createCell(3);
                    Cell PVCtxt = row1.createCell(4);
                    Cell GSWtxt = row1.createCell(5);
                    Cell Processtxt = row1.createCell(6);
                    Cell Machinetxt = row1.createCell(7);
                    Cell Aluminumtxt = row1.createCell(8);
                    this.Copper = rs.getDouble("Cu");
                    this.GSW = rs.getDouble("GSW");
                    this.PVC = rs.getDouble("PVC");
                    this.Aluminum = rs.getDouble("Aluminum");
                    String Process = rs.getString("Stage");
                    String Machines = rs.getString("Machine");
                    
                    TotalCopper = TotalCopper + Copper;
                    TotalPVC = TotalPVC + PVC;
                    TotalGSW = TotalGSW + GSW;
                    TotalAluminum = TotalAluminum + Aluminum;
           
                    Cutxt.setCellValue(Copper);
                    PVCtxt.setCellValue(PVC);
                    GSWtxt.setCellValue(GSW);
                    Aluminumtxt.setCellValue(Aluminum);
                    Processtxt.setCellValue(Process);
                    Machinetxt.setCellValue(Machines);
                    m++;
                    
                }
                
System.out.println(String.format("%.2f",TotalCopper) + " " + String.format("%.2f",TotalPVC) + ""
        + " " + String.format("%.2f",TotalGSW) + String.format("%.2f",TotalAluminum) +  Machiness[i] );
              
                 
                 XSSFRow row2  = actualreport.getRow(ss);
                 //Cell Machine = row2.createCell(11);
                 Cell Cu = row2.getCell(1);
                 Cell PVc = row2.getCell(4);
                 Cell GsW = row2.getCell(10);
                 Cell Al = row2.getCell(7);
                 //Machine.setCellValue(Machiness[i]);
                 Cu.setCellValue(Double.parseDouble(String.format("%.2f",TotalCopper)));
                 PVc.setCellValue(Double.parseDouble(String.format("%.2f",TotalPVC)));
                 GsW.setCellValue(Double.parseDouble(String.format("%.2f",TotalGSW)));
                 Al.setCellValue(Double.parseDouble(String.format("%.2f",TotalAluminum)));
                 ss++;
            }
            input.close();
            FileOutputStream outputStream = new FileOutputStream("SCrap101.xlsx");
            wb.write(outputStream);
            wb.close();
            outputStream.close();
        } catch (IOException | InvalidFormatException | EncryptedDocumentException | ClassNotFoundException | SQLException ee) {
            System.out.println(ee.getMessage());
        }
 
    }
    public void DeleteTable(){
        
     try{
      si.executeUpdate("DELETE FROM Output1");
     }catch(SQLException ww){
         System.out.println(ww.getMessage());
     }
    }

}
