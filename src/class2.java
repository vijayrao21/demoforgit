import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class class2 {
XSSFWorkbook wb;
XSSFSheet sh;
public class2(String excelpath) 
{
try {
	File f1=new File(excelpath);//creating object for file path
	FileInputStream fi=new FileInputStream(f1);//creating object to open file
	wb=new XSSFWorkbook(fi);//creating object for workbook
} catch (FileNotFoundException e) {//exception handling
		e.printStackTrace();
} catch (IOException e) {
		e.printStackTrace();
}}
public String getdatafromexcel(int sheetnum, int row,int col)
{
sh=wb.getSheetAt(sheetnum);//creating object for a sheet
String data=sh.getRow(row).getCell(col).getStringCellValue();//get data from a particular cell
return data;//return data from excel cell
}
int getlastrow(int sheetnum)
{
sh=wb.getSheetAt(sheetnum);
int lastrow=sh.getLastRowNum();//to get no of rows present in the excell
return lastrow;//returning last row
}}