package apac;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFExcel5 {
	
	public void Copypaste(int a,int b) throws IOException {
		File f=new File("../ApachePOI/excel009.xlsx");
		FileInputStream fi=new FileInputStream(f);
	    XSSFWorkbook xw=new XSSFWorkbook(fi);
	    XSSFSheet xs=xw.getSheetAt(0);
	    DataFormatter df=new DataFormatter();
        FormulaEvaluator fe=new XSSFFormulaEvaluator((XSSFWorkbook)xw);
        
        File f1=new File("../ApachePOI/excel012.xlsx");
		FileOutputStream fo=new FileOutputStream(f1);
	    XSSFWorkbook wk=new XSSFWorkbook();
	    XSSFSheet xt=wk.createSheet();
	    
	    int r=xs.getPhysicalNumberOfRows();
		for(int i=0;i<a;i++) {
			XSSFRow xr=xs.getRow(i);
			int c=xr.getPhysicalNumberOfCells();
			XSSFRow xr1=xt.createRow(i);
			for(int j=0;j<b;j++) {
				XSSFCell xc =xr.getCell(j);
				fe.evaluate(xc);
				String cellvalue=df.formatCellValue(xc,fe);
				
			 XSSFCell xc1=xr1.createCell(j);
						System.out.println(cellvalue);
						xc1.setCellValue(cellvalue);
												
			    }}wk.write(fo);
				fo.flush();
				fo.close();}
	
	public static void main(String[] args) throws IOException {
		XSSFExcel5 obj=new XSSFExcel5();
	Scanner s=new Scanner(System.in);
    System.out.println("Enter row");
    int p=s.nextInt();
    System.out.println("Enter column");
    int o=s.nextInt();
    
    obj.Copypaste(p,o);
	}

}
