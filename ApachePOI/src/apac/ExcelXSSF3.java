package apac;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelXSSF3 {//read data of row's
	
	public static void main(String[] args  ) throws IOException {
		File f=new File("../ApachePOI/excel009.xlsx");
         FileInputStream fi=new FileInputStream(f);
         XSSFWorkbook wk=new XSSFWorkbook(fi);
         XSSFSheet ws=wk.getSheetAt(0);
         Scanner s=new Scanner(System.in);
         System.out.println("Enter row 1");
         int a=s.nextInt();
         System.out.println("Enter row 2");
         int b=s.nextInt();
         DataFormatter df=new DataFormatter();
         FormulaEvaluator fe=new XSSFFormulaEvaluator((XSSFWorkbook)wk);
         int r=ws.getPhysicalNumberOfRows();
         for (int i=0;i<r;i++) {
        	 XSSFRow xr=ws.getRow(i);
        	 int c=xr.getPhysicalNumberOfCells();
        	 if((i>=a)&&(i<=b)) {
        	 for (int j=0;j<c;j++) {
        		 XSSFCell xc=xr.getCell(j);
        	 fe.evaluate(xc);
      String cellvalue= df.formatCellValue(xc, fe);
        	System.out.println("Cell "+j+i+" "+"Contains :"+xc.getCellType()+" :"+cellvalue);	 
        	 }
        	 
        	 
        	 
         }}
         
	}

}
