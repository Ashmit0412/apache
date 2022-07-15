package apac;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelXSSF2 {//read based on row
	
	public static void main(String[] args) throws IOException {
		File f=new File("../ApachePOI/excel008.xlsx");
		FileInputStream fi=new FileInputStream(f);
           XSSFWorkbook wk=new XSSFWorkbook(fi );
         XSSFSheet s1=wk.getSheetAt(0);
         int r=s1.getPhysicalNumberOfRows();
         Scanner s=new Scanner(System.in); 
 		System.out.println("enter Row #");
 		int l=s.nextInt();
 		
         for(int i=0;i<r;i++) {
        	 XSSFRow xr= s1.getRow(i);
        	 int c=xr.getPhysicalNumberOfCells();
        
        	 if (i==l) {
        		 for (int j=0;j<c;j++) 
        		{
        		 XSSFCell xc=xr.getCell(j);
        		 System.out.println(xc.getCellType());
        		System.out.println(xc.getStringCellValue());
        	 	 }
        		 } }
         	}}
