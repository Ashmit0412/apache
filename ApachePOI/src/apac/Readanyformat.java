package apac;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readanyformat {

	public static void main(String[] args) throws IOException {
		File f=new File("../ApachePOI/excel009.xlsx");
		FileInputStream fi=new FileInputStream(f);
           XSSFWorkbook wk=new XSSFWorkbook(fi );
         XSSFSheet s1=wk.getSheetAt(0);
         int r=s1.getPhysicalNumberOfRows();
         Scanner s=new Scanner(System.in); 
         DataFormatter objDefaultFormat = new DataFormatter();
 		FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) wk);

 		
         for(int i=0;i<r;i++) {
        	 XSSFRow xr= s1.getRow(i);
        	 int c=xr.getPhysicalNumberOfCells();
        
        	  {
        		 for (int j=0;j<c;j++) 
        		{
        		 XSSFCell xc=xr.getCell(j);
        		
        		 objFormulaEvaluator.evaluate(xc); // This will evaluate the cell, And any type of cell will return string value
     		    String cellValueStr = objDefaultFormat.formatCellValue(xc,objFormulaEvaluator);

        		System.out.println("Cell "+j+i+" contains :" +xc.getCellType()+":  " + cellValueStr);
        	 	 }
        		 } }}}
         	
		   
		

	
