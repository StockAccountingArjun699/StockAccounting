package Utilities;

import java.io.FileInputStream;

import java.util.Properties;



public class ReadingData 
{
	public static void main(String[] args) throws Throwable
	{
		
		Properties configproperties=new Properties();
		FileInputStream fis=new FileInputStream
		("C:\\Users\\nagarjuna.y\\workspace\\StockAccounting\\MavenStockAccounting\\PropertyFiles\\Environment.properties");
		configproperties.load(fis);
		System.out.println(configproperties.getProperty("Browser"));
		
		System.out.println(configproperties.getProperty("URL"));
		
		System.out.println(configproperties.getProperty("UserName"));
		System.out.println(configproperties.getProperty("Password"));
		
		
		
	}
}
