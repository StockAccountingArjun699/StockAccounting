package Utilities;

import java.io.FileInputStream;

import java.util.Properties;

public class PropertyFileUtil
{

	public static String getValueForKey(String key) throws Throwable
	{
		Properties configProperties=new Properties();
		
		FileInputStream fis=new FileInputStream
	
	("C:\\Users\\nagarjuna.y\\workspace\\StockAccounting\\MavenStockAccounting\\PropertyFiles\\Environment.properties");
		
		configProperties.load(fis);
		
		return configProperties.getProperty(key);
	}
}
