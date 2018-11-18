package spreadSheetFormat;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class RunFormatter {
	
	
 private String inputFile;
 private  Workbook w;
 private WritableWorkbook workbook;
 private WritableSheet sheet;
 
 
 private String currentServiceFunction;
 private String currentRelationship;
 private String currentInformationSubject;
 private String lastServiceFunction;


 private String relationshipToWrite = "";
 private String informationSubjectToWrite = "";
 
 
 private int count = 0;
 
 
 public void setFiles(String inputFile) throws IOException {
	 //Import file
	 this.inputFile = inputFile;
  
  	//Export file Original PC
  	//String fileName = "C:\\Users\\Alex\\Documents\\Internship\\New Excel file.xls";
	 
	//Export file Mac
	 String fileName = "/Users/alex/Documents/Internship/New Excel file.xls";
	 
	 
	// Creates a new work book
	workbook = Workbook.createWorkbook(new File(fileName));
	// Creates a new sheet on the workbook
	sheet = workbook.createSheet("Sheet1", 0);
 }
 

	public void writeToFile(String aCurrentServiceFunction, String aRelationshipToWrite,
			String aInformationSubjectToWrite) throws WriteException, IOException 
	{

		//MAKE SURE TO GET RIGHT COLUMN AND ROW
		// Creates a label column1
		Label label = new Label(0, count, aCurrentServiceFunction);
		// Creates a label column2
		Label label1 = new Label(1, count, aRelationshipToWrite);
		// Creates a label column3
		Label label2 = new Label(2, count, aInformationSubjectToWrite);
		// Creates a number
		//Number number = new Number(0, 1, 3.14159265358979323846);

		// Adds the labels
		sheet.addCell(label);
		sheet.addCell(label1);
		sheet.addCell(label2);
		// Adds a Number to the first row second column on sheet1
		//sheet.addCell(number);

		
	}

public void write() throws WriteException, IOException{
 
	File inputWorkbook = new File(inputFile);
	 
	 try {
	  w = Workbook.getWorkbook(inputWorkbook);
	  // Get the first sheet
	  Sheet inputsheet = w.getSheet(0);
	  
	  //First row not part of sort
	  
	  
	  //initialize the last variables so they can be compared
	  //GET EXACT COLUMN
	  Cell cellColumn1 = inputsheet.getCell(0,0);
	  Cell cellColumn2 = inputsheet.getCell(1,0);
	  Cell cellColumn3 = inputsheet.getCell(2,0);
	  writeToFile(cellColumn1.getContents(), cellColumn2.getContents(), cellColumn3.getContents());
	  count++;
	  
	  //initialize a compare
	  cellColumn1 = inputsheet.getCell(0,1);
	  lastServiceFunction = cellColumn1.getContents();
	  
	  //for(int j= 0;j<inputsheet.getColumns();j++){
	   for (int i=1;i<inputsheet.getRows(); i++) {
		   cellColumn1 = inputsheet.getCell(0,i);
		   cellColumn2 = inputsheet.getCell(1,i);
		   cellColumn3 = inputsheet.getCell(2,i);
		
		  	//Column 1
			currentServiceFunction = cellColumn1.getContents();
			//Column 2
			currentRelationship = cellColumn2.getContents();
			//Column 3
			currentInformationSubject = cellColumn3.getContents(); 
			if (currentServiceFunction.equals(lastServiceFunction))
			{
				relationshipToWrite += currentRelationship + ", ";
				
				// add currentInformationSubject to what's already in informationSubjctToWrite
				informationSubjectToWrite += currentInformationSubject + ", ";
			}	
			else
			{
				// remove last ", " from relationshipToWrite and informationSubjectToWrite
				//WRITE A FUNCTION!!!!!!
				relationshipToWrite = relationshipToWrite.substring(0,relationshipToWrite.length()-2);
				informationSubjectToWrite = informationSubjectToWrite.substring(0,informationSubjectToWrite.length()-2);

				// write currentServiceFunction and relationshipToWrite and informationSubjectToWrite to target file	
				//CALL THE WRITE FUNTION!!!!!
				writeToFile(currentServiceFunction, relationshipToWrite, informationSubjectToWrite);
				count++;
				
				relationshipToWrite = "";
				informationSubjectToWrite = "";
				relationshipToWrite += currentRelationship + ", ";

				// add currentInformationSubject to what's already in informationSubjctToWrite
				informationSubjectToWrite += currentInformationSubject + ", ";
			}
			
			
	    
	    
	    
	    //No need for a number
	    /*if(type == CellType.NUMBER) {
	     System.out.println("I got a number " + cell.getContents());
	    }
	    */
	   //}
	  }
	   workbook.write();
	  } catch(BiffException e) {
	   e.printStackTrace();
	 }
	 w.close();
	workbook.close();
	 
 
//	for(int j= 0;j<sheet.getColumns();j++)
//	{
//		for (int i=0;i<sheet.getRows(); i++) 
//		{
//			Cell cellColumn1 = sheet.getCell(j,i);
//			Cell cellColumn2 = sheet.getCell(j+1,i);
//			Cell cellColumn3 = sheet.getCell(j+2,i);
//			//Column 1
//			currentServiceFunction = cellColumn1.getContents();
//			//Column 2
//			currentRelationship = cellColumn2.getContents();
//			//Column 3
//			currentInformationSubject = cellColumn3.getContents(); 
//			if (currentServiceFunction == lastServiceFunction)
//			{
//				relationshipToWrite += currentRelationship + ", ";
//				
//				// add currentInformationSubject to what's already in informationSubjctToWrite
//				informationSubjectToWrite += currentInformationSubject + ", ";
//			}	
//			else
//			{
//				// remove last ", " from relationshipToWrite and informationSubjectToWrite
//				//WRITE A FUNCTION!!!!!!
//				relationshipToWrite.remove();
//				informationSubjectToWrite.remove();
//
//				// write currentServiceFunction and relationshipToWrite and informationSubjectToWrite to target file	
//				//CALL THE WRITE FUNTION!!!!!
//				writeToFile(currentServiceFunction, relationshipToWrite, informationSubjectToWrite);
//				
//				relationshipToWrite = "";
//				informationSubjectToWrite = "";
//				relationshipToWrite += currentRelationship + ", ";
//
//				// add currentInformationSubject to what's already in informationSubjctToWrite
//				informationSubjectToWrite += currentInformationSubject + ", ";
//			}
//		}
//	}
 }

	public static void main(String[] args) throws IOException, WriteException {
		RunFormatter test = new RunFormatter();
		//PC
		//test.setFiles("C:\\Users\\Alex\\Documents\\Internship\\retail_create-use_relationships for BG Import.xls");
		//Mac
		test.setFiles("/Users/alex/Documents/Internship/Foo.xls");
		test.write();
	}

}
