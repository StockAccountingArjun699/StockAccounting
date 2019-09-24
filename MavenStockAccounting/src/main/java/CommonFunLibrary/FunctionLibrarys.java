package CommonFunLibrary;

import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import org.openqa.selenium.interactions.Actions;



public class FunctionLibrarys 
{
	WebDriver driver;
	String res;
	//Launching Stock Accounting appLication
	public String appLaunch(String url)
	{
		System.setProperty("webdriver.chrome.driver", 
		" C:\\Users\\nagarjuna.y\\Downloads\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.get(url);
		driver.manage().window().maximize();
		//validation
		if (driver.findElement(By.id("username")).isDisplayed()) 
		{
			res="Pass";
		}else
		{
			res="Fail";
		}
		
		return res;
	}
	
	//Login into stockaccounting
	
	public String appLogin(String username,String password)
	{
		driver.findElement(By.id("username")).clear();
		driver.findElement(By.id("username")).sendKeys(username);
		driver.findElement(By.id("password")).clear();
		driver.findElement(By.id("password")).sendKeys(password);
		driver.findElement(By.id("btnsubmit")).click();
		//validation
		if (driver.findElement(By.id("logout")).isDisplayed()) 
		{
			res="Pass";
		}else
		{
			res="Fail";
		}
		
		return res;	
		
	}
	// Supplier Creation
	
	public String supplierCreation(String sName,String add,String city,String country,String cPerson,String pNumber,String email,String mNumber,String note) throws Throwable
	{
		driver.findElement(By.id("mi_a_suppliers")).click();
		driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[3]/div[1]/div[1]/div[1]/div/a")).click();
		Thread.sleep(2000);
		String exp_data=driver.findElement(By.id("x_Supplier_Number")).getAttribute("value");
		System.out.println(exp_data);
		driver.findElement(By.id("x_Supplier_Name")).sendKeys(sName);
		driver.findElement(By.name("x_Address")).sendKeys(add);
		driver.findElement(By.id("x_City")).sendKeys(city);
		driver.findElement(By.id("x_Country")).sendKeys(country);
		driver.findElement(By.id("x_Contact_Person")).sendKeys(cPerson);
		driver.findElement(By.name("x_Phone_Number")).sendKeys(pNumber);
		driver.findElement(By.id("x__Email")).sendKeys(email);
		driver.findElement(By.id("x_Mobile_Number")).sendKeys(mNumber);
		driver.findElement(By.id("x_Notes")).sendKeys(note);
		//scroll down page
		/*Actions pageDown=new Actions(driver);
		pageDown.sendKeys(Keys.PAGE_DOWN).build().perform();*/
		driver.findElement(By.id("btnAction")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//*[text()='OK!']")).click();
		Thread.sleep(4000);
		driver.findElement(By.xpath("//*[@class='ajs-button btn btn-primary']")).click();
		//validation
		if(driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[2]/div[2]/div/button")).isDisplayed())
		{
			driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[2]/div[2]/div/button")).click();
			Thread.sleep(3000);
			driver.findElement(By.id("psearch")).clear();
			driver.findElement(By.id("psearch")).sendKeys(exp_data);
			driver.findElement(By.id("btnsubmit")).click();
		}else
		{
			Thread.sleep(3000);
			driver.findElement(By.id("psearch")).clear();
			driver.findElement(By.id("psearch")).sendKeys(exp_data);
			driver.findElement(By.id("btnsubmit")).click();
		}
		
		String act_data=driver.findElement(By.xpath("//*[@id='el1_a_suppliers_Supplier_Number']/span")).getText();
		
		System.out.println(act_data);
		if (exp_data.equals(act_data)) 
		{
			res="Pass";
		}else
		{
			res="Fail";
		}
		driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[2]/div[2]/div/button")).click();
		
		return res;
	}
	//Stock Categories
	public String stockCategories(String cName) throws Exception
	{
		Actions mouse=new Actions(driver);
		mouse.moveToElement(driver.findElement(By.id("mi_a_stock_items"))).build().perform();
	mouse.moveToElement(driver.findElement(By.id("mi_a_stock_categories"))).click().build().perform();
	
	driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[3]/div[1]/div[1]/div[1]/div/a")).click();
	driver.findElement(By.id("x_Category_Name")).sendKeys(cName);
	
	driver.findElement(By.id("btnAction")).click();
	Thread.sleep(4000);
	driver.findElement(By.xpath("//*[text()='OK!']")).click();
	Thread.sleep(4000);
	driver.findElement(By.xpath("(//*[text()='OK'])[6]")).click();
	
	//validation
	if (driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[2]/div[2]/div/button")).isDisplayed())
	{
		driver.findElement(By.xpath("//*[@id='ewContentColumn']/div[2]/div[2]/div/button")).click();
		driver.findElement(By.id("psearch")).sendKeys(cName);
		driver.findElement(By.id("btnsubmit")).click();
	}else
	{
		driver.findElement(By.id("psearch")).sendKeys(cName);
		driver.findElement(By.id("btnsubmit")).click();
	}
	Thread.sleep(4000);
		String act_data=driver.findElement(By.xpath("//*[@id='el1_a_stock_categories_Category_Name']/span")).getText();
	System.out.println(act_data);
	if(cName.equals(act_data))
	{
		res="Pass";
	}else
	{
		res="Fail";
	}
	return res;
	}
	//Logout from StockAccounting
	
	public String appLogout() throws Throwable
	{
		driver.findElement(By.id("logout")).click();
		Thread.sleep(200);
		driver.findElement(By.xpath("//*[text()='OK!']")).click();
		//validation
				if (driver.findElement(By.id("username")).isDisplayed()) 
				{
					res="Pass";
				}else
				{
					res="Fail";
				}
				
				return res;
	}
	
	//close the StockAccounting
	
	public void appClose()
	{
		driver.close();
	}
	
}
