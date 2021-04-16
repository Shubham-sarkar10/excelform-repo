package toolsQa;

import java.awt.AWTException;
import org.openqa.selenium.support.ui.Select;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.ElementClickInterceptedException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class ExcelForm 
{

	public static void main(String[] args) throws IOException, Exception 
	
	{
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\shubham.sarkar\\Downloads\\chromedriver_win32\\chromedriver.exe");
		
		
		File file=new File("C:\\Users\\shubham.sarkar\\Downloads\\attachments\\Book1.xls");//FILE PROPERTIES PATH...
		  
			FileInputStream fileInput = null;
			
				fileInput = new FileInputStream(file);
			
				HSSFWorkbook obj=new HSSFWorkbook(fileInput);
				
				
				HSSFSheet obj1=obj.getSheet("Sheet1");
				
				
		WebDriver obj2=new ChromeDriver();
		obj2.get("https://demoqa.com/automation-practice-form");
		obj2.manage().window().maximize();
		
			WebElement firstName=obj2.findElement(By.id("firstName"));
	        WebElement lastName=obj2.findElement(By.id("lastName"));
	        WebElement email=obj2.findElement(By.id("userEmail"));
	        WebElement genderMale= obj2.findElement(By.id("gender-radio-1"));
	        WebElement mobile=obj2.findElement(By.id("userNumber"));
	        WebElement dob= obj2.findElement(By.id("dateOfBirthInput"));
	        WebElement subjects= obj2.findElement(By.id("subjectsInput"));
	        WebElement hobbies= obj2.findElement((By.id("hobbies-checkbox-1")));
	        WebElement path=obj2.findElement(By.id("uploadPicture"));
	      //  WebElement state=obj2.findElement(By.id("state"));
	        
	        WebElement address=obj2.findElement(By.id("currentAddress"));
	        WebElement statevalue=obj2.findElement(By.id("react-select-3-input"));
	        WebElement submitBtn=obj2.findElement(By.id("submit"));
	 

			 int rowCount = obj1.getLastRowNum();
			    System.out.println(rowCount);
			    for (int i = 1; i <=rowCount ; i++) 
			    {
			    	    firstName.sendKeys(obj1.getRow(i).getCell(0).getStringCellValue());
			            lastName.sendKeys(obj1.getRow(i).getCell(1).getStringCellValue());
			            email.sendKeys(obj1.getRow(i).getCell(2).getStringCellValue());
			           // genderMale.sendKeys(obj1.getRow(i).getCell(3).getStringCellValue());
			            mobile.sendKeys(obj1.getRow(i).getCell(4).getStringCellValue());
			            
			          //Click on the gender radio button using javascript
			            
			            JavascriptExecutor js = (JavascriptExecutor) obj2;
			            js.executeScript("arguments[0].click();", genderMale);
			            
			            //dob.sendKeys(args);
			    		//obj1.findElement((By.xpath("//option[contains(text(),'October')]"))).click();
			    		//obj1.findElement((By.xpath("//option[contains(text(),'1997')]"))).click();
			    		//obj1.findElement((By.xpath("//div[contains(text(),'20')]"))).click();
			    				
			            
			            subjects.sendKeys(obj1.getRow(i).getCell(5).getStringCellValue());
			            Robot robot=null;
			    		try {
			    		robot = new Robot();
			    		} catch (AWTException e) {
			    		e.printStackTrace();
			    		}
			    		
			    		robot.keyPress(KeyEvent.VK_DOWN);
			    		robot.keyRelease(KeyEvent.VK_DOWN);
			    		robot.keyPress(KeyEvent.VK_ENTER);
			    		robot.keyRelease(KeyEvent.VK_ENTER);
			    		
			    		
			    		JavascriptExecutor js1 = (JavascriptExecutor) obj2;
			            js1.executeScript("arguments[0].click();", hobbies);
			            
			    		//hobbies.sendKeys(obj1.getRow(i).getCell(6).getStringCellValue());
			    		//hobbies.click();
			    		
			    		path.sendKeys("C:\\Users\\shubham.sarkar\\Documents");
			    		//path.sendKeys(obj1.getRow(i).getCell(7).getStringCellValue());
			    		
			    		address.sendKeys(obj1.getRow(i).getCell(7).getStringCellValue());
			            
			    		statevalue.sendKeys(obj1.getRow(i).getCell(8).getStringCellValue());
			    		statevalue.click();
			    		
			    		
						
						 //JavascriptExecutor js2 = (JavascriptExecutor) obj2;
						// js2.executeScript("arguments[0].click();", state);
						  
						  //JavascriptExecutor js3 = (JavascriptExecutor) obj2;
						//  js3.executeScript("arguments[0].click();", statevalue);
						 
			    		 Thread.sleep(5000);
			           
			            //Click on submit button
			            submitBtn.click();
			            
			           
			            
			            obj2.findElement(By.id("firstName")).clear();
			            obj2.findElement(By.id("lastName")).clear();
			            obj2.findElement(By.id("userEmail")).clear();
			            obj2.findElement(By.id("gender-radio-1")).clear();
			            obj2.findElement(By.id("userNumber")).clear();
			            obj2.findElement(By.id("dateOfBirthInput")).clear();
			            obj2.findElement(By.id("subjectsInput")).clear();
			            obj2.findElement(By.id("hobbies-checkbox-1")).clear();
			            obj2.findElement(By.id("uploadPicture")).clear();
			           // obj2.findElement(By.id("state")).clear(); 
			            obj2.findElement(By.id("currentAddress")).clear();
			            obj2.findElement(By.id("react-select-3-input")).clear();
			            //obj2.findElement(By.id("state")).clear();
			    }
			            
			            
			          /*//Verify the confirmation message
			            WebElement confirmationMessage = obj2.findElement(By.xpath("//div[text()='Thanks for submitting the form']"));
			            
			            //create a new cell in the row at index 6
			            HSSFCell cell = obj1.getRow(i).createCell(6);
			            
			            //check if confirmation message is displayed
			            if (confirmationMessage.isDisplayed()) {
			                // if the message is displayed , write PASS in the excel sheet
			                cell.setCellValue("PASS");
			                
			            } else {
			                //if the message is not displayed , write FAIL in the excel sheet
			                cell.setCellValue("FAIL");
			            }
			            
			            // Write the data back in the Excel file
			            FileOutputStream outputStream = new FileOutputStream("C:\\Users\\shubham.sarkar\\Desktop\\Book1.xls");
			            obj.write(outputStream);
			 
			            //close the confirmation popup
			            WebElement closebtn = obj2.findElement(By.id("closeLargeModal"));
			            closebtn.click();
			            
			            //wait for page to come back to registration page after close button is clicked
			            obj2.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
			        }
			        
			        //Close the workbook
			        obj.close();
			        
			        //Quit the driver
			        obj2.quit();
			        }
			    	*/
			        
			        
			        
			            /*Row row = obj1.getRow(i);
			        // create a loop to print cell values

			        for (int j = 0; j < row.getLastCellNum(); j++) {
			            Cell cell = row.getCell(j);
			            
			            switch (cell.getCellType()) 
			            {
			            case STRING:
			                System.out.print(row.getCell(j).getStringCellValue() + " ");
			                break;

			            case NUMERIC:
			            	System.out.print(row.getCell(j).getNumericCellValue() + " ");
			                break;*/

			            
			           
			            
			        
			        //System.out.println();

			        

	   
		/* HSSFCell cell=row1.getCell(0);
		 HSSFCell cell1=row1.getCell(1);
		 String address= cell.getStringCellValue();
		 String address1= cell1.getStringCellValue();
		 System.out.println(address);
		 System.out.println(address1);
	}*/

	
	}
}


