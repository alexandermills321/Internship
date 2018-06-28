package spreadSheetFormat;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.Number;

public class RunFormatter {
 private String inputFile;
 private  Workbook w;
 public void setInputFile(String inputFile) {
  this.inputFile = inputFile;
 }
 
 public void read() throws IOException{
  File inputWorkbook = new File(inputFile);
 
 try {
  w = Workbook.getWorkbook(inputWorkbook);
  // Get the first sheet
  Sheet sheet = w.getSheet(0);
  //loop over first 10 column and lines
  
  for(int j= 0;j<sheet.getColumns();j++){
   for (int i=0;i<sheet.getRows(); i++) {
    Cell cell = sheet.getCell(j,i);
    CellType type = cell.getType();
    
    if(type == CellType.LABEL) {
     System.out.println("I got a label " + cell.getContents());
    }
    //No need for a number
    /*if(type == CellType.NUMBER) {
     System.out.println("I got a number " + cell.getContents());
    }
    */
   }
  }
  
  } catch(BiffException e) {
   e.printStackTrace();
 }
 


 }
 public void write() throws WriteException, IOException{
	 
		    String fileName = "C:\\Users\\Alex\\Documents\\Internship\\New Excel file.xls";
		    //Creates a new work book
		    WritableWorkbook workbook = Workbook.createWorkbook(new File(fileName));
		    //Creates a new sheet on the workbook
		    WritableSheet sheet1 = workbook.createSheet("Sheet1", 0);
		    
		    
		    //Creates a label
		    Label label = new Label(0,0,"We did it!!!! One label down 3000+ more to go.");
		    //Creates a number
		    Number number = new Number(0,1,3.14159265358979323846);
		    
		    //Adds a label to the first cell
		    sheet1.addCell(label);
		    //Adds a Number to the first row second column on sheet1
		    sheet1.addCell(number);
		    
		    workbook.write();
		    workbook.close();
 }
 
 public static void main(String[] args) throws IOException, WriteException {
  RunFormatter test= new RunFormatter();
  test.setInputFile("C:\\Users\\Alex\\Documents\\Internship\\retail_create-use_relationships for BG Import.xls");
  test.read();
  test.write();
 }
 
 
}

