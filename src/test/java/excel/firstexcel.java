package excel;

import org.testng.annotations.Test;
import java.io.BufferedReader;
import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashSet;
import java.util.Random;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import com.github.javafaker.Faker;
public class firstexcel {
		@Test
		public void m1() {
		try {
			 File status = new File("C:\\Users\\user\\project\\Excel\\mayank.xls");
             if (status.delete())
             {
                  System.out.println("File deleted successfully");
             }
             HSSFWorkbook workbook = new HSSFWorkbook();  
         	FileOutputStream fileOut = new FileOutputStream("C:\\Users\\user\\project\\Excel\\mayank.xls");  
         	
         	HSSFSheet sheet = workbook.createSheet("January"); 
         	HSSFRow rowhead = sheet.createRow((short)0);
         	rowhead.createCell(0).setCellValue("indexno");  
         	rowhead.createCell(1).setCellValue("Userid");  
         	rowhead.createCell(2).setCellValue("firstname");  
         	rowhead.createCell(3).setCellValue("last name");  
         	rowhead.createCell(4).setCellValue("username"); 
         	rowhead.createCell(5).setCellValue("password"); 
         	rowhead.createCell(6).setCellValue("emailid"); 
         	
         	
         		//fileOut.close();
         		//workbook.close();  
         	System.out.println("Created");
         	
         	
         	System.out.println("Enter the number of rows you want to enter");
           //	Scanner sc = new Scanner(System.in);
           //	int n = sc.nextInt();
           	Faker faker = new Faker();

           	for(int i=1;i<=100;i++) {
           		
           		String str1 = passwordgenerator();
           		
           		String indexno=indexno();
           		
           		String firstName = faker.name().firstName();
           		String userid = firstName.concat(indexno);
           		String lastName = faker.name().lastName();
           		String str3=firstName.concat(".");
           		String username=str3.concat(lastName);
           		String emailid = username.concat("@gmail.com");
           		HSSFRow row = sheet.createRow((short)i);
           		row.createCell(0).setCellValue(indexno);  
           		
           		row.createCell(1).setCellValue(userid);  
           		
           		row.createCell(2).setCellValue(firstName);  
           		row.createCell(3).setCellValue(lastName);  
           		row.createCell(4).setCellValue(username); 
           		row.createCell(5).setCellValue(str1); 
           		row.createCell(6).setCellValue(emailid); 
           		
           		
           		
           		
           	}
           	workbook.write(fileOut); 
         	fileOut.close();
         	  workbook.close();
           
		}
		catch (Exception e)   
		{  
		e.printStackTrace();  
		} 
		}
	

public static String passwordgenerator() {
	String lexicon = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "0123456789"
            + "abcdefghijklmnopqrstuvxyz" + "@#";
	
	Random rand = new Random();
	HashSet<String> identifiers = new HashSet<String>();

      
    StringBuilder builder = new StringBuilder();
    while(builder.toString().length() == 0) {
    	   int length = rand.nextInt(5)+5;
        for(int i = 0; i < length; i++) {
            builder.append(lexicon.charAt(rand.nextInt(lexicon.length())));
        }
        if(identifiers.contains(builder.toString())) {
            builder = new StringBuilder();
        }
    }
    return builder.toString();
}

public static String indexno() {
	String lexicon="123456789";
	
	Random rand = new Random();
	HashSet<String> identifiers = new HashSet<String>();

      
    StringBuilder builder = new StringBuilder();
  
    while(builder.toString().length() == 0) {
    	   int length = rand.nextInt(2)+2;
        for(int i = 0; i < length; i++) {
            builder.append(lexicon.charAt(rand.nextInt(lexicon.length())));
        }
        if(identifiers.contains(builder.toString())) {
            builder = new StringBuilder();
        }}
    // System.out.println(length);
    	   return builder.toString();
      }
	
}
