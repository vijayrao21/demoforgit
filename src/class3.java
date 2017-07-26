import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class class3 {
XSSFWorkbook wb;
XSSFSheet sh;
public class3(String excelpath){
try {
		File f1=new File(excelpath);
		FileInputStream fi=new FileInputStream(f1);
		wb=new XSSFWorkbook(fi);
	} catch (FileNotFoundException e) {
			e.printStackTrace();
	} catch (IOException e) {
			e.printStackTrace();
	}}
public void writeexcel(String excelpath, int sheetnum, int row, int col, String list1)
{
	try {
		sh=wb.getSheetAt(sheetnum);
		sh.createRow(row).createCell(col).setCellValue(list1);//set values for particlar cell
		FileOutputStream fout=new FileOutputStream(excelpath);//create object to pass values to file
		wb.write(fout);//write in the file
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}}}