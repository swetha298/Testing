package demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class excellsheet {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		 WebDriverWait wait;
		    HSSFWorkbook workbook;
		    HSSFSheet sheet;
		    HSSFCell cell;
		    
		    System.setProperty("WebDriver.chrome Driver","E:\\autom\\chromedriver_win32 (1)\\chromedriver.exe");	
			WebDriver driver=new ChromeDriver();
		    
		    
		//    driver.manage().window().maximize();

		    wait = new WebDriverWait(driver,30);
		    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
		
		
		
		 // Import excel sheet.
		     File src=new File("C:\\Users\\swetha\\Desktop\\ddt1.xls");
		      
		     // Load the file.
		     FileInputStream finput = new FileInputStream(src);
		      
		     // Load he workbook.
		    workbook = new HSSFWorkbook(finput);
		      
		     // Load the sheet in which data is stored.
		     sheet= workbook.getSheetAt(0);
		      
		     for(int i=1; i<=sheet.getLastRowNum(); i++)
		    	 
		     {
					driver.get("https://www.amazon.in/");

		         // Import data for iTEM.
		         cell = sheet.getRow(i).getCell(0);
		         cell.getStringCellValue();
		         driver.findElement(By.id("twotabsearchtextbox")).sendKeys(cell.getStringCellValue());
		         driver.findElement(By.id("nav-search-submit-button")).click();
		          
		         // Import data for password.
		      /*   cell = sheet.getRow(i).getCell(1);
		         cell.getStringCellValue();
		         driver.findElement(By.xpath("//input[@placeholder=\"Last Name\"]")).sendKeys(cell.getStringCellValue());
		          
		        */
		         // Write data in the excel.
		       FileOutputStream foutput=new FileOutputStream(src);
		         
		        // Specify the message needs to be written.
		        String message = "item found";
		        String message1 = "item not found";

		        if(driver.getTitle().contains(cell.getStringCellValue()))
		         {
		        // Create cell where data needs to be written.
		        sheet.getRow(i).createCell(1).setCellValue(message);
		         }
		        else
		        {
			        sheet.getRow(i).createCell(1).setCellValue(message1);

		        }
		        
		        // Specify the file in which data needs to be written.
		        FileOutputStream fileOutput = new FileOutputStream(src);
		         
		        // finally write content
		        workbook.write(fileOutput);
		         
		         // close the file
		        fileOutput.close();
		
		
		
		
		     }
		
		
	}

}
