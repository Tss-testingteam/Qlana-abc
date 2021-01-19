package Modules;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Contact {

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
	

		
		File file = new File("C:\\\\Users\\\\pc\\\\Desktop\\\\Contact.xlsx");   //creating a new file instance  
		FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
		//creating Workbook instance that refers to .xlsx file  
		XSSFWorkbook wb = new XSSFWorkbook(fis);   
		XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
		XSSFRow row = sheet.getRow(0);
		int colNum = row.getLastCellNum();

		System.out.println("Total Number of Columns in the excel is : " + colNum);
		int rowNum = sheet.getLastRowNum() + 1;
		System.out.println("Total Number of Rows in the excel is : " + rowNum);

		
		Row r;
		Cell c1, c2, c3, c4, c5, c6, c7;
		WebElement firstname = driver.findElementById("first_name");
		WebElement lastname = driver.findElementById("last_name");
		WebElement email = driver.findElementById("contact_info.email");
		WebElement phone = driver.findElementById("contact_info.phone_main");


		for (int i = 0; i <= rowNum; i++) {

			r = sheet.getRow(i);
			c1 = r.getCell(0);
			c2 = r.getCell(1);
			c3 = r.getCell(2);
			c4 = r.getCell(3);
			c5 = r.getCell(4);
			c6 = r.getCell(5);
			c7 = r.getCell(6);

			// DataFormatter formatter = new DataFormatter();
			// String str = formatter.formatCellValue(c4);
			// double dnum = Double.parseDouble(str);
			// System.out.println("In formated Cell Value--" + dnum);
			
			String prefixinp = c1.getStringCellValue();
			String firstnameinp = c2.getStringCellValue();
			String lastnameinp = c3.getStringCellValue();
			String emailinp = c4.getStringCellValue();
			String ccodeinp = c5.getStringCellValue();
			String phoneinp = c6.getStringCellValue();
			String contacttypeinp = c7.getStringCellValue();

			
		
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

			
			
			firstname.sendKeys(firstnameinp);
			Thread.sleep(1000);


			lastname.sendKeys(lastnameinp);
			Thread.sleep(1000);


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

			
	
			
			
//			savebtn.click();
//			Thread.sleep(2000);

		}
		
	
	}
}
