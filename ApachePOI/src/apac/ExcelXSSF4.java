package apac;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelXSSF4 {//write data give column and row and take data from user

	public void write1(int a, int b) throws IOException{
		Scanner s=new Scanner(System.in);
		File f=new File("../ApachePOI/excel010.xlsx");
		FileOutputStream fo=new FileOutputStream(f);
		XSSFWorkbook wk=new XSSFWorkbook();
		XSSFSheet xs=wk.createSheet();
		for (int i=0;i<a;i++) {
			XSSFRow xr=xs.createRow(i);
			for (int j=0;j<b;j++) {
				System.out.println("enter data for cell "+j+i);
				String k=s.nextLine();
				XSSFCell xc=xr.createCell(j);
				xc.setCellValue(k);
			}
		}wk.write(fo);
		fo.flush();
		fo.close();
		
	}
	
	public static void main(String[] args) throws IOException {
		Scanner s=new Scanner(System.in);
		
		System.out.println("Enter final row");
		int x=s.nextInt();
		System.out.println("Enter final column");
		int y=s.nextInt();
		ExcelXSSF4 obj=new ExcelXSSF4();
		obj.write1(x, y);
	}

}
