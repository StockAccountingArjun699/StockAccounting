package DriverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunLibrary.FunctionLibrary;
import Utilities.ExcelFileUtil;

public class DriverScript
{

	ExtentReports report;
	ExtentTest logger;
	WebDriver driver;
  
	public void startTest() throws Throwable
    {
    	ExcelFileUtil excel=new ExcelFileUtil();
    	
    	//MasterTestCases Sheet
    	for ( int i = 1;i<=excel.rowCount("MasterTestCases");i++)
    	{
    		String ModuleStatus="";
    		
    		if(excel.getData("MasterTestCases",i,2).equalsIgnoreCase("Y"))
    		{
    			//Define Module Name
    			String TCModule=excel.getData("MasterTestCases", i, 1);
    			report=new ExtentReports("D:\\nagarjuna\\workspace\\stockaccounting\\Reports\\"+TCModule+FunctionLibrary.generateDate()+".html");
    			 logger = report.startTest(TCModule);
    			
    			int rowcount=excel.rowCount(TCModule);
    			//TCModule sheet
    			for (int j=1; j<=rowcount;j++)
    				
    			{
    				String Description=excel.getData(TCModule, j, 0);
    				String Object_Type=excel.getData(TCModule, j, 1);
    				String Locator_Type=excel.getData(TCModule, j, 2);
    				String Locator_Value=excel.getData(TCModule, j, 3);
    				String Test_Data=excel.getData(TCModule, j, 4);
    				try
    				{
    				if(Object_Type.equalsIgnoreCase("StartBrowser"))
    				{
    					driver=FunctionLibrary.startBrowser(driver);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("openApplication"))
    				{
    					FunctionLibrary.openApplication(driver);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("typeAction"))
    				{
    					FunctionLibrary.typeAction(driver,Locator_Type,Locator_Value,Test_Data);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("clickAction"))
    				{
    					FunctionLibrary.clickAction(driver,Locator_Type,Locator_Value);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("closeBrowser"))
    				{
    					FunctionLibrary.closeBrowser(driver);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("waitForElement"))
    				{
    					FunctionLibrary.waitForElement(driver,Locator_Type,Locator_Value,Test_Data);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("pageDown"))
    				{
    					FunctionLibrary.pageDown(driver);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("captureData"))
    				{
    					FunctionLibrary.captureData(driver, Locator_Type, Locator_Value);
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("tableValidation"))
    				{
    					FunctionLibrary.tableValidation(driver,Test_Data );
    					logger.log(LogStatus.INFO, Description);
    				}
    				if(Object_Type.equalsIgnoreCase("mouseClick"))
    				{
    			
    					logger.log(LogStatus.INFO, Description);
    				}
    					
    				if(Object_Type.equalsIgnoreCase("waitForElement"))
    				excel.setData(TCModule, j, 5, "pass");
    				logger.log(LogStatus.PASS, Description+"		PASS ");
    				ModuleStatus="true";
    				}catch(Exception e)
    				{
    					//Status update in TCModule Sheet "fail"
    					excel.setData(TCModule, j, 5, "FAIl");
    					logger.log(LogStatus.FAIL,Description+"   FAIL");
    					ModuleStatus="false";
    					File srcFile=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
    				
FileUtils.copyFile(srcFile,new File("C:\\Users\\nagarjuna.y\\workspace\\StockAccounting\\MavenStockAccounting\\Screenshots"+TCModule+"   "+FunctionLibrary.generateDate()+".png"));
    					break;
    				}
    			}
    			if(ModuleStatus.equalsIgnoreCase("true"))
    			{
    				//Status update in MasterTestCases Sheet "pass"
    				excel.setData("MasterTestCases", i, 3, "PASS");
    			}else
    				if(ModuleStatus.equalsIgnoreCase("false"))
    				{
    					//Statusupdate in MasterTestCases Sheet "fail"
    					excel.setData("MasterTestCases", i, 3, "FAIL");
    				}
    			report.endTest(logger);
    			report.flush();
    		}
    		else
    		{
    			excel.setData("MasterTestCases", i, 3, "Not Executed");
    		}
    		
    		
    	}
    }
}
	
