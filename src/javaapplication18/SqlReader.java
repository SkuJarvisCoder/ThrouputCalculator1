/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication18;

/**
 *
 * @author paul
 */
import java.sql.*;
public class SqlReader {
    Connection con;
    Statement st;
    ResultSet rs;
    String name, CuMass, PVCMass, BeddingMass, GSWMass, SheathMass, cores;
    double drawing;
    
    public void Reader(String item){
        String itemcode = item;
        if(item.length()<9){
           this.name = item;
            this.drawing  = DrawingMass(item); 
        }
        String url = "jdbc:ucanaccess://MatmassSiku1.accdb";
        String query = "SELECT * FROM Standards  WHERE  ItemCode ='" + itemcode + "' ";
       try{
        con = DriverManager.getConnection(url);
        st = con.createStatement();
        rs = st.executeQuery(query);
        while(rs.next()){
             this.name = rs.getString("Description");
              this.CuMass = rs.getString("Stranding");
              this.PVCMass = rs.getString("Insulation");
              this.BeddingMass = rs.getString("Bedding");
              this.GSWMass  = rs.getString("Armouring");
             this.SheathMass = rs.getString("Sheathing");
             this.cores = rs.getString("NoOfCores");
           
            
       }
        con.close();
    }catch(SQLException ee){
        System.out.println(ee.getLocalizedMessage()); 
    }
       
}
     public static double DrawingMass(String item){
         int density;
         String wires;
         String dia;
         if(item.length()==4){
             density = 2700;
              wires = "1";
              dia = item;
return ((Double.parseDouble(wires)*density*Math.PI*Math.pow((Double.parseDouble(dia)*0.001), 2))/4)*1000;      
         }
         else 
             density = 8890;
           int n = item.indexOf("/");
         wires = item.substring(0, n);
         dia = item.substring(n+1);
return ((Double.parseDouble(wires)*density*Math.PI*Math.pow((Double.parseDouble(dia)*0.001), 2))/4)*1000;
     }
    
    public double[] Amounts(String Process, double length){
        double[] stards = new double[3];
        double GSW=0;
         double PVC = 0;
         double CuOrAL = 0;
        switch(Process){
            case "Wire Drawing" :
                GSW = 0;
                PVC = 0;
                CuOrAL = length * drawing;
            break;
            case "Bunching" :
                GSW = 0;
                PVC = 0;
                CuOrAL = (length * Double.parseDouble(CuMass))/Double.parseDouble(cores);
            break;
            case "Stranding":
                GSW = 0;
                PVC = 0;
                CuOrAL = (length * Double.parseDouble(CuMass))/Double.parseDouble(cores);
            break;
            case "Insulation" :
                GSW = 0;
                PVC =  (length * Double.parseDouble(PVCMass))/Double.parseDouble(cores);
                CuOrAL = (length * Double.parseDouble(CuMass))/Double.parseDouble(cores);
            break;
            case "LayingUp" :
                GSW = 0;
                PVC = length * Double.parseDouble(PVCMass);
                   CuOrAL = length * Double.parseDouble(CuMass);
            break;
            case "Bedding" :
                GSW = 0;
                PVC = length * Double.parseDouble(PVCMass) + (length * Double.parseDouble(BeddingMass));
                CuOrAL = length * Double.parseDouble(CuMass);
            break;
            case "Armouring" :
                GSW = length * Double.parseDouble(GSWMass);
                  PVC = length * Double.parseDouble(PVCMass) + (length * Double.parseDouble(BeddingMass));
                CuOrAL = length * Double.parseDouble(CuMass);
            break;
            case "Sheathing" :
             GSW = length * Double.parseDouble(GSWMass);
                  PVC = length * Double.parseDouble(PVCMass) + (length * Double.parseDouble(BeddingMass)) + (length * Double.parseDouble(SheathMass));
                CuOrAL = length * Double.parseDouble(CuMass);  
             break;
            case "Twinning" :
                GSW = 0;
                PVC =  ((length * Double.parseDouble(PVCMass))/Double.parseDouble(cores))*2;
                CuOrAL = ((length * Double.parseDouble(CuMass))/Double.parseDouble(cores))*2;
        }
        stards[0] = CuOrAL;
        stards[1] = PVC/1.02;
        stards[2] = GSW;
        return stards;
    }
   
} 
