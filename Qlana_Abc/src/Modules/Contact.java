package Modules;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import Xlutils.Xluitls;

public class Contact extends Xluitls {
	
	public static void main(String[] args)
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\pc\\Desktop\\Selenium Jars\\chromedriver_win32\\chromedriver.exe");
		ChromeDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://qlana.dev.zeroco.de/");
		Thread.sleep(2000);
		
		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
		driver.findElementById("appUserName").sendKeys("maria.a@bamboocp.com");
		Thread.sleep(3000);
		driver.findElementById("appPassword").sendKeys("The@1234");
		Thread.sleep(2000);
		driver.findElementById("loginBtn").click();
		Thread.sleep(2000);
		

		List<WebElement> moduleslist = 	driver.findElements(By.id("menu-list"));
		moduleslist.get(0).click();
	
		String xlfile = "C:\\Users\\pc\\Desktop\\Contact.xlsx";
		String xlsheet = "Sheet1";

		int rc = Xluitls.getRowCount(xlfile, xlsheet);
		
		File file = new File("C:\\\\Users\\\\pc\\\\Desktop\\\\Contact.xlsx");   //creating a new file instance  
		FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
		//creating Workbook instance that refers to .xlsx file  
		@SuppressWarnings("resource")
		XSSFWorkbook wb = new XSSFWorkbook(fis);   
		XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
		XSSFRow row = sheet.getRow(0);
		int colNum = row.getLastCellNum();

		System.out.println("Total Number of Columns in the excel is : " + colNum);
		int rowNum = sheet.getLastRowNum() + 1;
		System.out.println("Total Number of Rows in the excel is : " + rowNum);


		WebElement firstname = driver.findElementById("first_name");
		WebElement lastname = driver.findElementById("last_name");
		WebElement email = driver.findElementById("contact_info.email");
		WebElement phone = driver.findElementById("contact_info.phone_main");
     	WebElement savebtn = driver.findElementById("button_Save");



		for (int i = 0; i <= rc; i++) {

			String prefixinp = Xluitls.getCellData(xlfile, xlsheet, i, 0);
			String firstnameinp = Xluitls.getCellData(xlfile, xlsheet, i, 1);
			String lastnameinp = Xluitls.getCellData(xlfile, xlsheet, i, 2);
			String emailinp = Xluitls.getCellData(xlfile, xlsheet, i, 3);
			String ccodeinp = Xluitls.getCellData(xlfile, xlsheet, i, 4);
			String phoneinp = Xluitls.getCellData(xlfile, xlsheet, i, 5);
			String contacttypeinp = Xluitls.getCellData(xlfile, xlsheet, i, 6);

		
			driver.findElement(By.xpath("//*[@id=\"formly_9_select_salutation_0\"]/div/label")).click();
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").clear();
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(prefixinp);
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(Keys.ENTER);
			Thread.sleep(1000);

			
			firstname.clear();
			firstname.sendKeys(firstnameinp);
			Thread.sleep(1000);

			lastname.clear();
			lastname.sendKeys(lastnameinp);
			Thread.sleep(1000);

			email.clear();
			email.sendKeys(emailinp);
			Thread.sleep(1000);
			
			
			driver.findElement(By.xpath("//*[@id=\"formly_9_select_contact_info.country_main_4\"]/div/label")).click();
			Thread.sleep(1000);
			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").clear();
			Thread.sleep(1000);
			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(ccodeinp);
			Thread.sleep(1000);
			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(1000);
			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(Keys.ENTER);
			Thread.sleep(1000);
			
			phone.clear();
			phone.sendKeys(phoneinp);
			Thread.sleep(1000);
			
			driver.findElement(By.xpath("//*[@id=\"formly_9_select_contact_type_6\"]/div/label")).click();
			Thread.sleep(1000);
			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").clear();
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(contacttypeinp);
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(1000);

			driver.findElementByXPath("//input[@class='ui-dropdown-filter ui-inputtext ui-widget ui-state-default ui-corner-all']").sendKeys(Keys.ENTER);
			Thread.sleep(1000);

			savebtn.click();
			Thread.sleep(2000);
			
			String msg = driver.findElementByClassName("toast-message").getText();
			System.out.println(msg);
			
			

			if(msg.equalsIgnoreCase("Successfully saved"))
			{
				
	   	            //  System.out.println("Success msg displayed");
	   	              Xluitls.setCellData(xlfile, xlsheet, i, 7, msg);
	   	              Xluitls.setCellData(xlfile, xlsheet, i, 8, "Pass");
					  Xluitls.fillGreenColor(xlfile, xlsheet, i, 8);
			
				
			}else if(msg.equalsIgnoreCase("Email already exist")) {
				
		            System.out.println("Failed due to Email already exist");
		           Xluitls.setCellData(xlfile, xlsheet, i, 7, msg);
	   	           Xluitls.setCellData(xlfile, xlsheet, i, 8, "Fail");
				   Xluitls.fillRedColor(xlfile, xlsheet, i, 8);
		           
			}else if(msg.equalsIgnoreCase("Given phone no. already exists")) {				
		           //System.out.println("Null data");
		           Xluitls.setCellData(xlfile, xlsheet, i, 7, "Given phone no. already exists");
	   	           Xluitls.setCellData(xlfile, xlsheet, i, 8, "Fail");
				   Xluitls.fillRedColor(xlfile, xlsheet, i, 8);
			}else if(msg.equalsIgnoreCase(" ")) {	
				
				 //System.out.println("Null data");
		           Xluitls.setCellData(xlfile, xlsheet, i, 7, "Mandatory field is missing or Invalid data is given");
	   	           Xluitls.setCellData(xlfile, xlsheet, i, 8, "Fail");
				   Xluitls.fillRedColor(xlfile, xlsheet, i, 8);
					

		}
		
		}
	}
}
