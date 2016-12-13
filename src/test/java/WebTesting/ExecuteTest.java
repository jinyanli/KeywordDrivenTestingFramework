package WebTesting;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;


import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Listeners;

/**
 * THIS IS THE EXAMPLE OF KEYWORD DRIVEN TEST CASE
 *
 */
public class ExecuteTest {
	private WebDriver driver;
	private WebDriverWait myWaitVar;

	
    @AfterTest
	public void testLogin() throws Exception {
		// TODO Auto-generated method stub
        
        System.setProperty( "webdriver.chrome.driver", "C:\\Users\\1323928\\Desktop\\selenium\\chromedriver.exe" );
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        
        Properties allObjects = new Properties();

        //Read object repository file
        InputStream stream = new FileInputStream(new File(System.getProperty("user.dir")+"\\src\\test\\java\\objects\\object.properties"));
        //load all objects
        allObjects.load(stream);
        
        //Read keyword sheet
        //Sheet mySheet = readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "KeywordFramework");
        Sheet mySheet = readExcel(System.getProperty("user.dir")+"\\","TestCase.xlsx" , "Sheet2");
        //Find number of rows in excel file
    	//int rowCount = mySheet.getLastRowNum()-mySheet.getFirstRowNum();
    	//Create a loop over all the rows of excel file to read it
    	for (int i = 1; i <= mySheet.getLastRowNum(); i++) {
    		//Loop over all the rows
    		Row row = mySheet.getRow(i);
    		//Check if the first cell contain a value, if yes, That means it is the new testcase name
    		if(row.getCell(0).toString().length()==0){
    		//Print testcase detail on console
    			System.out.println(row.getCell(1).toString()+"----"+ row.getCell(2).toString()+"----"+
    			row.getCell(3).toString()+"----"+ row.getCell(4).toString());
    		//Call perform function to perform operation on UI
    			perform(driver, allObjects, row.getCell(1).toString(), row.getCell(2).toString(),
    				row.getCell(3).toString(), row.getCell(4).toString());
    	    }
    		else{
    			//Print the new  testcase name when it started
    				System.out.println("New Testcase->"+row.getCell(0).toString() +" Started");
    			}
    	}
	}
    
    public Sheet readExcel(String filePath,String fileName,String sheetName) throws IOException{
    	//Create a object of File class to open xlsx file
    	File file =	new File(filePath+"\\"+fileName);
    	//Create an object of FileInputStream class to read excel file
    	FileInputStream inputStream = new FileInputStream(file);
    	Workbook myWorkbook = null;
    	//Find the file extension by spliting file name in substing and getting only extension name
    	String fileExtensionName = fileName.substring(fileName.indexOf("."));
    	//Check condition if the file is xlsx file
    	if(fileExtensionName.equals(".xlsx")){
    	//If it is xlsx file then create object of XSSFWorkbook class
    		myWorkbook = new XSSFWorkbook(inputStream);
    	}
    	//Check condition if the file is xls file
    	else if(fileExtensionName.equals(".xls")){
    		//If it is xls file then create object of XSSFWorkbook class
    		myWorkbook = new HSSFWorkbook(inputStream);
    	}
    	//Read sheet inside the workbook by its name
    	Sheet  mySheet = myWorkbook.getSheet(sheetName);
    	 return mySheet;	
    }
    
    public void perform(WebDriver wd, Properties p,String operation,String objectName,String objectType,String value) throws Exception{
        System.out.println("");
        switch (operation.toUpperCase()) {
        case "CLICK":
            //Perform click
            driver.findElement(getObject(p,objectName,objectType)).click();
            break;
        case "SETTEXT":
            //Set text on control
            driver.findElement(getObject(p,objectName,objectType)).sendKeys(value);
            break;
            
        case "GOTOURL":
            //Get url of application
            driver.get(p.getProperty(value));
            break;
        case "GETTEXT":
            //Get text of an element
            driver.findElement(getObject(p,objectName,objectType)).getText();
            break;
        case "ClickIfPresent":
            //Get text of an element
        	if(driver.findElements(getObject(p,objectName,objectType)).size()!=0)
        		driver.findElement(getObject(p,objectName,objectType)).click();
            break;
        default:
            break;
        }
    }

    
    private By getObject(Properties p,String objectName,String objectType) throws Exception{
        //Find by xpath
        if(objectType.equalsIgnoreCase("XPATH")){
            
            return By.xpath(p.getProperty(objectName));
        }
        //find by class
        else if(objectType.equalsIgnoreCase("CLASSNAME")){
            
            return By.className(p.getProperty(objectName));
            
        }
        //find by name
        else if(objectType.equalsIgnoreCase("NAME")){
            
            return By.name(p.getProperty(objectName));
            
        }
        //Find by css
        else if(objectType.equalsIgnoreCase("CSS")){
            
            return By.cssSelector(p.getProperty(objectName));
            
        }
        //find by link
        else if(objectType.equalsIgnoreCase("LINK")){
            
            return By.linkText(p.getProperty(objectName));
            
        }
        //find by partial link
        else if(objectType.equalsIgnoreCase("PARTIALLINK")){
            
            return By.partialLinkText(p.getProperty(objectName));
            
        }else if(objectType.equalsIgnoreCase("ID")){
            
            return By.id(p.getProperty(objectName));
            
        }else  	
        {
            throw new Exception("Wrong object type");
        }
    }
}



