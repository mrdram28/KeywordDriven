package testcases;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class ExecuteTest {
	
	private WebDriver driver;       
    @Test              
    public void testEasy() {  
    	
    	System.out.println("#####################################");
    	
    	System.out.println("Starting Firefox Driver");
    	
        driver.get("http://www.guru99.com/selenium-tutorial.html"); 
        
        System.out.println("Opened Guru Url");
        
        String title = driver.getTitle();   
        
        System.out.println("Title is : "+title);
        
        Assert.assertTrue(title.contains("Free Selenium Tutorials"));
        
        System.out.println("#####################################");
    }   
    @BeforeTest
    public void beforeTest() {  
        driver = new FirefoxDriver();  
    }       
    @AfterTest
    public void afterTest() {
        driver.quit();          
    }       

}
