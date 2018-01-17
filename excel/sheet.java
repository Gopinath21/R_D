import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class OrderFesiblitySheet {
		
	public static int LOCIncrement=1;
	//public static int col;
	public static int noRows=500;
	public static XSSFRow newRow;
	public static XSSFCell newCell;
	
	public static void main(String[] args) throws Exception {

		FileOutputStream fos=new FileOutputStream(new File("NBN_Order_Feasiblity.xlsx"));
		XSSFWorkbook NOFS=new XSSFWorkbook();
		XSSFSheet sheet	=	NOFS.createSheet();
		XSSFRow		headerrow=		sheet.createRow(0);
		XSSFCell cell00	=headerrow.createCell(0);
						cell00.setCellType(CellType.STRING);
						cell00.setCellValue("Location_ID");
						
		XSSFCell		cell01		=headerrow.createCell(1);
						cell01.setCellType(CellType.STRING);
						cell01.setCellValue("Product_Type");
						
		XSSFCell   cell02=headerrow.createCell(2);
					cell02.setCellType(CellType.STRING);
					cell02.setCellValue("Scenario_Type");
					
		XSSFCell    cell03=headerrow.createCell(3);
					cell03.setCellType(CellType.STRING);
					cell03.setCellValue("InteractionStatus");
					
		XSSFCell    cell04=headerrow.createCell(4);
					cell04.setCellType(CellType.STRING);
					cell04.setCellValue("NTD Shortfall");
					
		XSSFCell   cell05=headerrow.createCell(5);
					cell05.setCellType(CellType.STRING);
					cell05.setCellValue("LEADIN Shortfall");
					
		XSSFCell	cell06=headerrow.createCell(6);
					cell06.setCellType(CellType.STRING);
					cell06.setCellValue("NBNCOINFRASTRUCTURE shortfall");
					
		XSSFCell			cell07=headerrow.createCell(7);
							cell07.setCellType(CellType.STRING);
							cell07.setCellValue("Type");	
	
		XSSFCell			cell08=headerrow.createCell(8);	
							cell08.setCellType(CellType.STRING);
							cell08.setCellValue("Description");
							
		XSSFCell			cell09=headerrow.createCell(9);	
							cell09.setCellType(CellType.STRING);
							cell09.setCellValue("ExceptionType");
							
		XSSFCell			cell10=headerrow.createCell(10);	
							cell10.setCellType(CellType.STRING);
							cell10.setCellValue("ErrorCode");
					
		XSSFCell			cell11=headerrow.createCell(11);	
							cell11.setCellType(CellType.STRING);
							cell11.setCellValue("ErrorMessage");
		
							
	
							
		XSSFCell			cell12=headerrow.createCell(12);	
							cell12.setCellType(CellType.STRING);
							cell12.setCellValue("POWERSUPPLYWITHBATTERYBACKUP");
							
		XSSFCell          cell13=headerrow.createCell(13);	
		                  cell13.setCellType(CellType.STRING);
		                  cell13.setCellValue("BATTERBACKUPINSTALLDATE");
		                  
		XSSFCell          cell14=headerrow.createCell(14);
	    				  cell14.setCellType(CellType.STRING);
                          cell14.setCellValue("MIGRATION_FLAG");
       
       XSSFCell          cell15=headerrow.createCell(15);
	    				  cell15.setCellType(CellType.STRING);
                        cell15.setCellValue("PATCH");		
                        
                        
                        for(int row=1;row<=noRows;row++){
                           	newRow=sheet.createRow(row);
                           	
                           	for(int col=0;col<16;col++){
                           			if(col==0){
                           				newCell=newRow.createCell(col);
                           				newCell.setCellType(CellType.STRING);
                           				newCell.setCellValue("LOC"+LOCIncrement++);
                           			}
                           			if(col==1){
                           				newCell=newRow.createCell(col);
                           				newCell.setCellType(CellType.STRING);
                           				newCell.setCellValue("NCAS");
                           			}
                           			if(col==2){
                           				newCell=newRow.createCell(col);
                           				newCell.setCellType(CellType.STRING);
                           				newCell.setCellValue("NCAS-Shortfall");
                           			}if(col==3){
                           				newCell=newRow.createCell(col);
                           				newCell.setCellType(CellType.STRING);
                           				newCell.setCellValue("Feasible-Appointment Required");
                           			}if(col==4){
                           				newCell=newRow.createCell(col);
                           				newCell.setCellType(CellType.STRING);
                           				newCell.setCellValue("Yes");
                           			}if(col==5){
                           				newCell=newRow.createCell(col);
                           				newCell.setCellType(CellType.STRING);
                           				newCell.setCellValue("No");
                           				}if(col==6){
                           					newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.STRING);
                               				newCell.setCellValue("No");
                           				}if(col==7){
                           					newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.STRING);
                               				newCell.setCellValue("Standard install");
                           					
                           				}if(col==8){
                           					newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.BLANK);
                               				
                           					
                                          }if(col==9){
                                        	  newCell=newRow.createCell(col);
                                 				newCell.setCellType(CellType.BLANK);
                                 				
                                          }if(col==10){
                                        	  newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.BLANK);
                                        	  
                                          }if(col==11){
                                        	  newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.BLANK);
                                        	  
                                          }if(col==12){
                                        	  newCell=newRow.createCell(col);
                                 				newCell.setCellType(CellType.STRING);
                                 				newCell.setCellValue("No");
                                          }if(col==13){
                                        	  newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.BLANK);
                                        	  
                                          }if(col==14){
                                        	  newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.BLANK);
                                        	  
                                          }if(col==15){
                                        	  newCell=newRow.createCell(col);
                               				newCell.setCellType(CellType.BLANK);
                                        	  
                                          }
                                          
                           		
                           		
                           	}
                           	
                           		
             
                           	
                        	
                        }
                        
	
                        
                        
                        
                        
                        
                        NOFS.write(fos);
                        
                        
                        
	}
}
