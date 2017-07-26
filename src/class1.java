import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
public class class1 {
public static void main(String[] args) {
int t=0;
class2 a1=new class2("C:\\Users\\sony\\Desktop\\New Microsoft Office Excel Worksheet.xlsx");
class3 a2= new class3("C:\\Users\\sony\\Desktop\\New Microsoft Office Excel Worksheet.xlsx");
int lr=a1.getlastrow(0);
System.out.println("Number of rows in excel = "+lr);
ArrayList<String> list1=new ArrayList<String>();//create object for arraylist
for(int i=0;i<=lr;i++)
{
list1.add(a1.getdatafromexcel(0,i,0));//saving values in arraylist
}
System.out.println(list1);
for (int m=0;m<list1.size();m++){
a2.writeexcel("C:\\Users\\sony\\Desktop\\New Microsoft Office Excel Worksheet.xlsx",1,m,1,list1.get(m));
}
HashMap<String,String> hm=new HashMap<String,String>();  
for(int j=0;j<=lr;j++){
hm.put(a1.getdatafromexcel(0,j,0),a1.getdatafromexcel(0, j, 1)); } 
for(Map.Entry n:hm.entrySet()){  
	String p=(String) n.getKey();
	 String q=(String) n.getValue();
	a2.writeexcel("C:\\Users\\sony\\Desktop\\New Microsoft Office Excel Worksheet.xlsx",2,t,0,p);
	a2.writeexcel("C:\\Users\\sony\\Desktop\\New Microsoft Office Excel Worksheet.xlsx",2,t,2,q);
	 t=t+1;	
}}}