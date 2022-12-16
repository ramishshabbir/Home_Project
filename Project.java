package HomeProject;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.time.Duration;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Project {

	Logger log = Logger.getLogger("Project");
	public String BaseUrl = "https://galaxy.pk/";
	public WebDriver driver;
	public WebDriverWait wait;
	public Actions actions;
	public String Expected="" ,Actual="";


@BeforeTest
public void Main() {
	driver = new ChromeDriver();		
	driver.get(BaseUrl);
	PropertyConfigurator.configure("log4j.properties");
	driver.manage().window().maximize();
	
	//Create objects
	actions = new Actions(driver);
	wait = new WebDriverWait(driver, Duration.ofSeconds(5));						
}
@AfterTest
public void CloseUrl() {
	//driver.close();
}

//Quick Access Funtion
	public  WebElement WaitXpath(String xpath) {
		WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
		return element;
		}
	
	public void ActionsClick(WebElement element) {
		actions.click(element).perform();		
	}
	
	
	
	
@Test(priority=0)
public void HomePage() throws InterruptedException {
	//Hover on Products
	WebElement Products = WaitXpath("//ul[@id='menu-all-departments-1']//parent::div");
	actions.moveToElement(Products).perform();
	
	//Hover on Laptops and then click 
	WebElement Laptops = WaitXpath("//ul[@id='menu-all-departments-1']//child::li[@id='menu-item-4761']");
	actions.moveToElement(Laptops).perform();
	Thread.sleep(3000);
	ActionsClick(Laptops);
	
	//Back to Home
	WebElement Home = WaitXpath("//a[text()='Home']");
	ActionsClick(Home);
	log.info("Back to home");
}

@Test(priority=0)
public void Tablets() throws Exception  {
	
	 WebElement  description = WaitXpath("(//div[@class='product-loop-footer product-item__footer']//child::span[@class='woocommerce-Price-amount amount'])[\"+i+\"]");
	 WebElement  name = WaitXpath("(//div[@class='product-loop-body product-item__body']//child::h2[@class='woocommerce-loop-product__title'])[\"+i+\"]");
	 WebElement  price = WaitXpath("(//div[@class='product-loop-footer product-item__footer']//child::span[@class='woocommerce-Price-amount amount'])[\"+i+\"]");
	    		
	 for (int i=1; i<10; i++)
		{
	
	WebElement  image = WaitXpath("(//div[@class='product-thumbnail product-item__thumbnail'])[\"+i+\"]");
	image.click();
	String descs=description.getText();
	String namee= name.getText();
	String Price= price.getText();
	
	
	Actions right_click=new Actions(driver);
	right_click.contextClick(image).build().perform();
	Robot robot = new Robot();
	for(int n=0; n<7;n++)
	{
	robot.keyPress(KeyEvent.VK_DOWN);
	robot.keyRelease(KeyEvent.VK_DOWN);
	
	}
	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);
	Thread.sleep(2000);
	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);
	int j=1;
	//for(int j=1; j<3; j++)
	//{
	ExcelFn(namee, Price, descs,1,1);
		
		
	//}
	driver.navigate().back();
	Thread.sleep(3000);
	
	
		}	
		
}

public void ExcelFn( String namee ,String Price, String descs, int Row, int Col) throws Exception {

	XSSFWorkbook workbook1=new XSSFWorkbook();
	XSSFSheet sheet1=workbook1.createSheet("Computer Details");
	sheet1.createRow(Row).createCell(Col);
	sheet1.getRow(Row).createCell(0).setCellValue(Row);
	sheet1.getRow(Row).getCell(Col).setCellValue(namee);
	sheet1.getRow(Row).createCell(Col+1).setCellValue(Price);
	sheet1.getRow(Row).createCell(Col+2).setCellValue(descs);
	sheet1.createRow(0).createCell(0).setCellValue("Sr#");
	sheet1.getRow(0).createCell(1).setCellValue("Name");
	sheet1.getRow(0).createCell(2).setCellValue("Price");
	sheet1.getRow(0).createCell(3).setCellValue("Description");

	File fil=new File("C:\\Users\\4426\\eclipse-workspace\\Home_Project\\ExcelSheet.xlsx");		
	FileOutputStream fos=new FileOutputStream(fil);
	workbook1.write(fos);
		
	}






}
