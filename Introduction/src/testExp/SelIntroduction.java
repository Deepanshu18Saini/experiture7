package testExp;
import org.testng.annotations.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;


public class SelIntroduction {
	
	@Test
	public static void main() throws InterruptedException, Exception {
		// TODO Auto-generated method stub
		
		
		System.setProperty("webdriver.chrome.driver","C:\\Users\\Sandipan\\eclipse-workspace\\Drivers\\chromedriver_win32\\chromedriver.exe");
		
		//Opening Chrome in Incognito window
		ChromeOptions options=new ChromeOptions();
		options.addArguments("--incognito");
		WebDriver driver=new ChromeDriver(options);
		driver.manage().window().maximize();
		
		
		/*
		
		//ChromeDriver driver=new ChromeDriver();
		WebDriver driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://admin-portal.experiture.com/");
		
		*/
		
		StopWatch pageLoad = new StopWatch();
        pageLoad.start();
        driver.get("https://admin-portal.experiture.com/");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(120, 1));
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("btnLogin")));

        
        
        
        //Get the time
        long pageLoadTime_ms = pageLoad.getTime();
        long pageLoadTime_Seconds = pageLoadTime_ms / 1000;
        System.out.println(driver.getTitle());
        
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds + " seconds");
        
        pageLoad.stop();
        
        
        
        
        
        WebElement username=driver.findElement(By.cssSelector("#txtEmail")); 
        WebElement password=driver.findElement(By.cssSelector("#txtPassword")); 
        WebElement login=driver.findElement(By.xpath("//input[@id='btnLogin']")); 
        username.sendKeys("sandipanb@experiture.com"); password.sendKeys("Mapps123#"); 
        pageLoad.reset();
        pageLoad.start();
        login.click();
                
        driver.switchTo().frame("_iframe-dhtmlOpenAgencyClient");
        
       
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("txtClient")));

        
        //Get the time
        long pageLoadTime_ms1 = pageLoad.getTime();
        long pageLoadTime_Seconds1 = pageLoadTime_ms1 / 1000;
        System.out.println("Agency Selection");
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms1 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds1 + " seconds");
       // sheet.getRow(2).getCell(1).setCellValue(pageLoadTime_Seconds1);
        pageLoad.stop();
        
        
        //driver.findElement(By.id("txtClient")).sendKeys("Sycuan Casino Resort & Singing Hills Golf Resort");
        driver.findElement(By.id("txtClient")).sendKeys("New QA Team");

        driver.findElement(By.id("btnSearch")).click();
        pageLoad.reset();
        
        pageLoad.start();
        driver.findElement(By.id("obtClientList_ctl00__0")).click();
        
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#myheader_lblClientNameExp")));

        
        //Get the time
        long pageLoadTime_ms2 = pageLoad.getTime();
        long pageLoadTime_Seconds2 = pageLoadTime_ms2 / 1000;
        System.out.println(driver.getTitle());
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms2 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds2 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
        
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
				"/html/body[@id='pagebody']/form[@id='form1']/header[@id='exp_header']/div[@class='exptopheader']/a[@id='myheader_burgerMenu']/div[@id='menu-toggle']/div[@id='cross']")))
				.click();
       
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@id='aSTTargets']/i[@class='icon-right-thin-chevron']"))).click();
        
        // record time for Data Source and List page 
        pageLoad.start();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@id='SourcePgs']/a"))).click();
        
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'Manual')]")));

        
        //Get the time
        long pageLoadTime_ms3 = pageLoad.getTime();
        long pageLoadTime_Seconds3 = pageLoadTime_ms3 / 1000;
        System.out.println(driver.getTitle());
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms3 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds3 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        // Record time for All Targets Page
        pageLoad.start();
        driver.findElement(By.xpath("//span[contains(text(),'All Targets')]")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#imgBtnAddTargeti")));

        
        
      //Get the time
        long pageLoadTime_ms4 = pageLoad.getTime();
        long pageLoadTime_Seconds4 = pageLoadTime_ms4 / 1000;
        System.out.println("All Targets");
       //System.out.println("Total Page Load Time: " + pageLoadTime_ms4 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds4 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
     // Record time for Segments Page
        pageLoad.start();
        driver.findElement(By.xpath("//span[contains(text(),'Segments')]")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#imgBtnCreateView")));

        
        
      //Get the time
        long pageLoadTime_ms5 = pageLoad.getTime();
        long pageLoadTime_Seconds5 = pageLoadTime_ms5 / 1000;
        System.out.println("Segments");
       // System.out.println("Total Page Load Time: " + pageLoadTime_ms5 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds5 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
     // Record time for Conditional Segments Page
        pageLoad.start();
        driver.findElement(By.xpath("//a[@id='aConditionalSegment']")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Segments')]")));

        
        
      //Get the time
        long pageLoadTime_ms6 = pageLoad.getTime();
        long pageLoadTime_Seconds6 = pageLoadTime_ms6 / 1000;
        System.out.println("Conditional Segments");
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms6 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds6 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
     // Record time for Set Score Page
        pageLoad.start();
        driver.findElement(By.xpath("//span[contains(text(),'Set Score')]")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[@id='btnCalculateProfileScore']")));

        
        
      //Get the time
        long pageLoadTime_ms7 = pageLoad.getTime();
        long pageLoadTime_Seconds7 = pageLoadTime_ms7 / 1000;
        System.out.println("Set Score");
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms7 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds7 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
        
     // Record time for Subscriptions Page
        pageLoad.start();
        driver.findElement(By.xpath("//span[contains(text(),'Subscriptions')]")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[@id='imgBtnAddTargeti']")));

        
        
      //Get the time
        long pageLoadTime_ms8 = pageLoad.getTime();
        long pageLoadTime_Seconds8 = pageLoadTime_ms8 / 1000;
        System.out.println(driver.getTitle());
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms8 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds8 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
     // Record time for Suppression List Page
        pageLoad.start();
        driver.findElement(By.xpath("//a[@id='aSuppressList']")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Suppression Lists')]")));

        
        
      //Get the time
        long pageLoadTime_ms9 = pageLoad.getTime();
        long pageLoadTime_Seconds9 = pageLoadTime_ms9 / 1000;
        System.out.println(driver.getTitle());
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms9 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds9 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
        // Record time for Suppressed Domains Page
        pageLoad.start();
        driver.findElement(By.xpath("//a[@id='aSuppressDomain']")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Suppressed Domains')]")));

        
        
      //Get the time
        long pageLoadTime_ms10 = pageLoad.getTime();
        long pageLoadTime_Seconds10 = pageLoadTime_ms10 / 1000;
        System.out.println("Suppressed Domains");
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms10 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds10 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
        // Record time for  Data Model Page
        pageLoad.start();
        driver.findElement(By.xpath("//span[contains(text(),'Data Model')]")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Relational Tables')]")));

        
        
      //Get the time
        long pageLoadTime_ms11 = pageLoad.getTime();
        long pageLoadTime_Seconds11 = pageLoadTime_ms11 / 1000;
        System.out.println("Relational Tables");
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms11 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds11 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
     // Record time for  Relation Table Data Page
        pageLoad.start();
        driver.findElement(By.xpath("//a[@id='RelationalTableData']")).click();
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Relational Table Data')]")));

        
        
      //Get the time
        long pageLoadTime_ms12 = pageLoad.getTime();
        long pageLoadTime_Seconds12 = pageLoadTime_ms12 / 1000;
        System.out.println("Relational Tables Data");
        //System.out.println("Total Page Load Time: " + pageLoadTime_ms12 + " milliseconds");
        System.out.println("Total Page Load Time: " + pageLoadTime_Seconds12 + " seconds");
        pageLoad.stop();
        pageLoad.reset();
        
        
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
				"/html/body[@id='pagebody']/form[@id='form1']/header[@id='exp_header']/div[@class='exptopheader']/a[@id='myheader_burgerMenu']/div[@id='menu-toggle']/div[@id='cross']")))
				.click();
       
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@id='ProgramName']/i[@class='icon-right-thin-chevron  subO']"))).click();
        
        
     // Record time for  loading a Campaign
        pageLoad.start();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//body[1]/form[1]/header[1]/div[2]/div[3]/div[1]/div[5]/div[1]/ul[1]/li[3]/a[1]"))).click();
        
        
        wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#divOutbound")));

        
        
        //Get the time
          long pageLoadTime_ms13 = pageLoad.getTime();
          long pageLoadTime_Seconds13 = pageLoadTime_ms13 / 1000;
          System.out.println(driver.getTitle());
          //System.out.println("Total Page Load Time: " + pageLoadTime_ms13 + " milliseconds");
          System.out.println("Total Page Load Time: " + pageLoadTime_Seconds13 + " seconds");
          pageLoad.stop();
          pageLoad.reset();
        
       // Record time for  loading Analytics
          pageLoad.start();
          driver.findElement(By.xpath("//*[@id=\"aMTReports\"]/span")).click();
        
          wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#ctl00_ContentPlaceHolder1_btnViewLiveReprot")));
        
        
        
        //Get the time
          long pageLoadTime_ms14 = pageLoad.getTime();
          long pageLoadTime_Seconds14 = pageLoadTime_ms14 / 1000;
          System.out.println(driver.getTitle());
          //System.out.println("Total Page Load Time: " + pageLoadTime_ms14 + " milliseconds");
          System.out.println("Total Page Load Time: " + pageLoadTime_Seconds14 + " seconds");
          pageLoad.stop();
          pageLoad.reset();
        
        
          wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
  				"/html[1]/body[1]/form[1]/div[3]/div[5]/header[1]/a[1]/div[1]/div[2]"))).click();
         
          
       // Record time for  loading Dashboard
          pageLoad.start();
          wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/form[1]/div[3]/div[7]/div[1]/ul[1]/li[1]/a[1]"))).click();
        
        
          wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h3[@id='tourMP']")));
          
          
          
          //Get the time
            long pageLoadTime_ms15 = pageLoad.getTime();
            long pageLoadTime_Seconds15 = pageLoadTime_ms15 / 1000;
            System.out.println("Reports to Dashboard");
            //System.out.println("Total Page Load Time: " + pageLoadTime_ms15 + " milliseconds");
            System.out.println("Total Page Load Time: " + pageLoadTime_Seconds15 + " seconds");
            pageLoad.stop();
            pageLoad.reset();
        
         // Record time for  loading Reports from Dashboard
            pageLoad.start();
            driver.findElement(By.
            cssSelector("body.experiture:nth-child(2) div.dashboard:nth-child(47) div.lstContainer.row div.col-6:nth-child(1) div.box:nth-child(2) div.inner-content section.grpnl.rcntprg:nth-child(1) div.dashprgdata span:nth-child(1) span.row-style:nth-child(1) div.dtrow a.rpt > i.icon-progress-report")).click();
            
        
            wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#ctl00_ContentPlaceHolder1_btnViewLiveReprot")));        
        
          //Get the time
            long pageLoadTime_ms16 = pageLoad.getTime();
            long pageLoadTime_Seconds16 = pageLoadTime_ms16 / 1000;
            System.out.println("Dashboard to Reports");
            //System.out.println("Total Page Load Time: " + pageLoadTime_ms16 + " milliseconds");
            System.out.println("Total Page Load Time: " + pageLoadTime_Seconds16 + " seconds");
            pageLoad.stop();
            pageLoad.reset();
        
            
            
          //  wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
      		//		"/html[1]/body[1]/form[1]/div[3]/div[5]/header[1]/a[1]/div[1]/div[2]"))).click();
           
           // wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/form[1]/div[3]/div[7]/div[1]/ul[1]/li[5]/a[1]/i[2]"))).click();
           // Record time for  loading Landing Pages
              pageLoad.start();
             // wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#ctl00_ucHeader_aLeadPgs"))).click();
              driver.get("https://page.experiture.com/LeadPageDashboard.aspx?maintab=true");
              wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//span[contains(text(),'Lead Pages')]")));        
              
              //Get the time
                long pageLoadTime_ms17 = pageLoad.getTime();
                long pageLoadTime_Seconds17 = pageLoadTime_ms17 / 1000;
                System.out.println("Active Lead Pages");
                //System.out.println("Total Page Load Time: " + pageLoadTime_ms17 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds17 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
        
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//h3[contains(text(),'Draft Lead Pages')]"))).click();
                // Record time for  loading Landing Pages
             
                
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(" //label[contains(text(),'Launched on:')]")));   
                
                
              //Get the time
                long pageLoadTime_ms18 = pageLoad.getTime();
                long pageLoadTime_Seconds18 = pageLoadTime_ms18 / 1000;
                System.out.println("Draft Lead Pages");
                //System.out.println("Total Page Load Time: " + pageLoadTime_ms18 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds18 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/form[1]/header[1]/div[2]/div[2]/ul[1]/li[4]/i[1]"))).click();
                
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Surveys Forms')]"))).click();

              
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Forms')]")));  
                
                
                //Get the time
                long pageLoadTime_ms19 = pageLoad.getTime();
                long pageLoadTime_Seconds19 = pageLoadTime_ms19 / 1000;
                System.out.println(driver.getTitle());
               //System.out.println("Total Page Load Time: " + pageLoadTime_ms19 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds19 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/div[3]/form[1]/header[1]/div[2]/div[2]/ul[1]/li[4]/i[1]"))).click();
                
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'All Inbound Leads')]"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("body.experiture:nth-child(2) div.lstContainer:nth-child(42) div.rightPn1 div.prgtrgtCont:nth-child(5) div:nth-child(2) div.gridTop:nth-child(1) div.gridFilterL.gridFilterText:nth-child(2) > strong:nth-child(1)")));  
                
                
                //Get the time
                long pageLoadTime_ms20 = pageLoad.getTime();
                long pageLoadTime_Seconds20 = pageLoadTime_ms20 / 1000;
                System.out.println("All Inbound Leads");
               // System.out.println("Total Page Load Time: " + pageLoadTime_ms20 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds20 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@id='InboundMarketing']/i[@class='icon-right-thin-chevron']"))).click();

                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Lead Analytics')]"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[contains(text(),'Lead Dashboard Reports')]")));  

              //Get the time
                long pageLoadTime_ms21 = pageLoad.getTime();
                long pageLoadTime_Seconds21 = pageLoadTime_ms21 / 1000;
                System.out.println(driver.getTitle());
                //System.out.println("Total Page Load Time: " + pageLoadTime_ms21 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds21 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/form[1]/header[1]/div[2]/div[2]/ul[1]/li[4]/i[1]"))).click();
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Website Tracking')]"))).click();
                
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[contains(text(),'Pixel Tracking Code')]")));  
                
              //Get the time
                long pageLoadTime_ms22 = pageLoad.getTime();
                long pageLoadTime_Seconds22 = pageLoadTime_ms22 / 1000;
                System.out.println(driver.getTitle());
               // System.out.println("Total Page Load Time: " + pageLoadTime_ms22 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds22 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(),'Templates')]"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h2[@id='spanSourceWithCount']")));  
                
              //Get the time
                long pageLoadTime_ms23 = pageLoad.getTime();
                long pageLoadTime_Seconds23 = pageLoadTime_ms23 / 1000;
                System.out.println(driver.getTitle());
               // System.out.println("Total Page Load Time: " + pageLoadTime_ms23 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds23 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(),'Asset Manager')]"))).click();
                
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[@id='lnkUpload']")));
                
                               
              //Get the time
                long pageLoadTime_ms24 = pageLoad.getTime();
                long pageLoadTime_Seconds24 = pageLoadTime_ms24 / 1000;
                System.out.println(driver.getTitle());
              //  System.out.println("Total Page Load Time: " + pageLoadTime_ms24 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds24 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@id='cross']"))).click();
                
                //wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@id='Personalization']/i[@class='icon-right-thin-chevron']\""))).click();
                
                pageLoad.start();
                //wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Global Variables')]"))).click();
                driver.get("https://personalize.experiture.com/AccountVariables.aspx");
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//body[1]/form[1]/div[3]/div[1]/div[1]/h2[1]")));
                
              //Get the time
                long pageLoadTime_ms25 = pageLoad.getTime();
                long pageLoadTime_Seconds25 = pageLoadTime_ms25 / 1000;
                System.out.println(driver.getTitle());
             //   System.out.println("Total Page Load Time: " + pageLoadTime_ms25 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds25 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#cross"))).click();
                
                //wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Relational Widgets')]"))).click();
                
                pageLoad.start();
                driver.get("https://personalize.experiture.com/HtmlWidget.aspx");
                //wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Relational Widgets"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//body[1]/form[1]/div[3]/div[1]/div[1]/h2[1]")));
                
              //Get the time
                long pageLoadTime_ms26 = pageLoad.getTime();
                long pageLoadTime_Seconds26 = pageLoadTime_ms26 / 1000;
                System.out.println(driver.getTitle());
             //   System.out.println("Total Page Load Time: " + pageLoadTime_ms26 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds26 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#cross"))).click();
               // wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Conditional Segments')]"))).click();
                
                pageLoad.start();
                //wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Conditional Segments"))).click();
                driver.get("https://target.experiture.com/RuleSegmentManager.aspx");
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//body[1]/form[1]/div[5]/div[2]/div[1]/div[1]/div[1]/h2[1]")));
                
                
              //Get the time
                long pageLoadTime_ms27 = pageLoad.getTime();
                long pageLoadTime_Seconds27 = pageLoadTime_ms27 / 1000;
                System.out.println(driver.getTitle());
             //   System.out.println("Total Page Load Time: " + pageLoadTime_ms27 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds27 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#cross"))).click();
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/form[1]/header[1]/div[2]/div[2]/ul[1]/li[10]/a[1]"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//body[1]/form[1]/div[5]/div[1]/div[1]/div[3]/label[1]")));
                
              //Get the time
                long pageLoadTime_ms28 = pageLoad.getTime();
                long pageLoadTime_Seconds28 = pageLoadTime_ms28 / 1000;
                System.out.println(driver.getTitle());
             //   System.out.println("Total Page Load Time: " + pageLoadTime_ms28 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds28 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/form[1]/header[1]/div[2]/div[4]/div[1]/ul[1]/li[2]/a[1]"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#btnAddUser")));
                
              //Get the time
                long pageLoadTime_ms29 = pageLoad.getTime();
                long pageLoadTime_Seconds29 = pageLoadTime_ms29 / 1000;
                System.out.println(driver.getTitle());
              //  System.out.println("Total Page Load Time: " + pageLoadTime_ms29 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds29 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                
                
                pageLoad.start();
                wait.until(ExpectedConditions.elementToBeClickable(By.tagName("figure"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("#tourMP")));
                
              //Get the time
                long pageLoadTime_ms30 = pageLoad.getTime();
                long pageLoadTime_Seconds30 = pageLoadTime_ms30 / 1000;
                System.out.println("Accounts to Dashboard");
             //   System.out.println("Total Page Load Time: " + pageLoadTime_ms30 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds30 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
               //trying to logout 
                driver.findElement(By.xpath("/html[1]/body[1]/form[1]/header[1]/div[1]/div[1]/div[2]/div[1]/i[1]")).click();
                
                pageLoad.start(); 
                wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Logout"))).click();
                
                wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//h1[contains(text(),'Sign in to Experiture')]")));
                
              //Get the time
                long pageLoadTime_ms31 = pageLoad.getTime();
                long pageLoadTime_Seconds31 = pageLoadTime_ms31 / 1000;
                System.out.println("Logout");
             //   System.out.println("Total Page Load Time: " + pageLoadTime_ms31 + " milliseconds");
                System.out.println("Total Page Load Time: " + pageLoadTime_Seconds31 + " seconds");
                pageLoad.stop();
                pageLoad.reset();
                
                String s0=String.valueOf(pageLoadTime_Seconds);
                String s1=String.valueOf(pageLoadTime_Seconds1);
                String s2=String.valueOf(pageLoadTime_Seconds2);
                String s3=String.valueOf(pageLoadTime_Seconds3);
                String s4=String.valueOf(pageLoadTime_Seconds4);
                String s5=String.valueOf(pageLoadTime_Seconds5);
                String s6=String.valueOf(pageLoadTime_Seconds6);
                String s7=String.valueOf(pageLoadTime_Seconds7);
                String s8=String.valueOf(pageLoadTime_Seconds8);
                String s9=String.valueOf(pageLoadTime_Seconds9);
                String s10=String.valueOf(pageLoadTime_Seconds10);
                String s11=String.valueOf(pageLoadTime_Seconds11);
                String s12=String.valueOf(pageLoadTime_Seconds12);
                String s13=String.valueOf(pageLoadTime_Seconds13);
                String s14=String.valueOf(pageLoadTime_Seconds14);
                String s15=String.valueOf(pageLoadTime_Seconds15);
                String s16=String.valueOf(pageLoadTime_Seconds16);
                String s17=String.valueOf(pageLoadTime_Seconds17);
                String s18=String.valueOf(pageLoadTime_Seconds18);
                String s19=String.valueOf(pageLoadTime_Seconds19);
                String s20=String.valueOf(pageLoadTime_Seconds20);
                String s21=String.valueOf(pageLoadTime_Seconds21);
                String s22=String.valueOf(pageLoadTime_Seconds22);
                String s23=String.valueOf(pageLoadTime_Seconds23);
                String s24=String.valueOf(pageLoadTime_Seconds24);
                String s25=String.valueOf(pageLoadTime_Seconds25);
                String s26=String.valueOf(pageLoadTime_Seconds26);
                String s27=String.valueOf(pageLoadTime_Seconds27);
                String s28=String.valueOf(pageLoadTime_Seconds28);
                String s29=String.valueOf(pageLoadTime_Seconds29);
                String s30=String.valueOf(pageLoadTime_Seconds30);
                String s31=String.valueOf(pageLoadTime_Seconds31);

                try (// workbook object
				XSSFWorkbook workbook = new XSSFWorkbook()) {
					// spreadsheet object
					XSSFSheet spreadsheet
					    = workbook.createSheet(" Page Load ");
        
					// creating a row object
					XSSFRow row;
        
					// This data needs to be written (Object[])
					Map<String, Object[]> studentData = new TreeMap<String, Object[]>();
        
					studentData.put("1",new Object[] { "Source Link Name", "Destination Link Name", "Page Load in Seconds" });
					
					studentData.put("2", new Object[] { "Start", "Experiture Login",s0});
        
					studentData.put("3", new Object[] { "Experiture Login", "Agency Selection", s1});
        
					studentData.put("4", new Object[] { "Agency Selection", "Account Dashboard",s2 });
        
					studentData.put("5", new Object[] { "Account Dashboard", "Data Sources & Lists",s3 });
        
					studentData.put("6", new Object[] { "Data Sources & Lists", "All Targets",s4 });
					
					studentData.put("7", new Object[] { "All Targets", "Segments",s5 });
					
					studentData.put("8", new Object[] { "Segments", "Conditional Segments",s6 });
					
					studentData.put("9", new Object[] { "Conditional Segments", "Set Score",s7 });
					
					studentData.put("10", new Object[] { "Set Score", "Unsubscribe",s8 });
					
					studentData.put("11", new Object[] { "Unsubscribe", "Suppressed Target",s9 });
					
					studentData.put("12", new Object[] { "Suppressed Target", "Suppressed Domains",s10 });
					
					studentData.put("13", new Object[] { "Suppressed Domains", "Relational Tables",s11 });
					
					studentData.put("14", new Object[] { "Relational Tables", "Relational Tables Data",s12 });
					
					studentData.put("15", new Object[] { "Relational Tables Data", "Programs",s13 });
					
					studentData.put("16", new Object[] { "Programs", "Experiture Analytics",s14 });
					
					studentData.put("17", new Object[] { "Experiture Analytics", "Dashboard",s15 });
					
					studentData.put("18", new Object[] { "Dashboard", "Reports",s16 });
					
					studentData.put("19", new Object[] { "Reports", "Active Lead Pages",s17 });
					
					studentData.put("20", new Object[] { "Active Lead Pages", "Draft Lead Pages",s18});
					
					studentData.put("21", new Object[] { "Draft Lead Pages", "Forms",s19 });
					
					studentData.put("22", new Object[] { "Forms", "All Inbound Leads",s20 });
					
					studentData.put("23", new Object[] { "All Inbound Leads", "Leads Dashboard",s21 });
					
					studentData.put("24", new Object[] { "Leads Dashboard", "Website Tracking",s22 });
					
					studentData.put("25", new Object[] { "Website Tracking", "Template Gallery",s23 });
					
					studentData.put("26", new Object[] { "Template Gallery", "Asset Manager",s24 });
					
					studentData.put("27", new Object[] { "Asset Manager", "Account Variables",s25 });
					
					studentData.put("28", new Object[] { "Account Variables", "HTML Widget",s26 });
					
					studentData.put("29", new Object[] { "HTML Widget", "Targets",s27 });
					
					studentData.put("30", new Object[] { "Targets", "Agency",s28 });
					
					studentData.put("31", new Object[] { "Agency", "Manage Users",s29 });
					
					studentData.put("32", new Object[] { "Manage Users", "Dashboard",s30 });
					
					studentData.put("33", new Object[] { "Dashboard", "Logout",s31 });
					
					
        
					Set<String> keyid = studentData.keySet();
        
					int rowid = 0;
        
					// writing the data into the sheets...
        
					for (String key : keyid) {
        
					    row = spreadsheet.createRow(rowid++);
					    Object[] objectArr = studentData.get(key);
					    int cellid = 0;
        
					    for (Object obj : objectArr) {
					        Cell cell = row.createCell(cellid++);
					        cell.setCellValue((String)obj);
					    }
					}


					Calendar cal=Calendar.getInstance();
					Date time=cal.getTime();
					String timestamp= time.toString().replace(" ", "").replace(":", "");
					
					
				
					// .xlsx is the format for Excel Sheets...
					// writing the workbook into the file...
					FileOutputStream out = new FileOutputStream(
					    new File("C:\\Users\\Sandipan\\Excel Test Data\\PageLoad_"+timestamp+".xlsx"));
        
					workbook.write(out);
					out.close();
				}
		
                System.out.println("Excel write successfull");
            
        		
		
                
                
                
                
                
                
                
                
                
                
	}
	
	
	}


