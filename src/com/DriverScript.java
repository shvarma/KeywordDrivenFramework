package com;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.data.JRXlsDataSource;
import net.sf.jasperreports.engine.util.JRLoader;

import org.apache.commons.io.FileUtils;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverBackedSelenium;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeDriverService;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxWebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;


import autoitx4java.AutoItX;

import com.jacob.com.LibraryLoader;
import com.thoughtworks.selenium.Selenium;

public class DriverScript  {
	
	private int indvar = 0;
	private int synctime = 1000;
	private String webdriver;
	private String baseurl;
	private String setsetting;
	private ResultSet rset, rset1, rset2;
	private Connection conn;
	private Statement stmt;	
	private WebDriver driver;
	private String Activity = null, Attrname = null, Attrvalue = null, Parameter = null, stepstatus = null, teststatus = null, StepDetails = null;
	private int Option1 = 0, Option2 = 0, Option3 = 0;
	private Selenium selenium;
	private Robot robot;
	private WebElement element;
	private Select select;
	private AutoItX aitxobj;
	private boolean ObjectExists;
	private ChromeDriverService service;
	private int nCols;
	private Actions builder;
	private String chromedriver;
	private int retrycount = 3;	
	private int recurcount = 1;	
	private String osarch = "amd64"; // x86 or amd64
	private String TestScript;
	
	@Test	
	public void testScripts() throws Exception {		
		
		openExcelDocument();
		String alertTitle = "Message from webpage";
		osarch = System.getProperty("os.arch");	
		String TestScripts = "Select * From [TestWare$] Where [execstatus] = TRUE;";
		rset1 = DBConnection(TestScripts);
		rset1.last();
		int NumOfTcs = rset1.getRow() - 1; // One default value in the excel sheet
		System.out.println("Testcases: " + NumOfTcs + " active testcase(s) in the suite");
		rset1.first();
		
		setsetting = rset.getString("testscript");
		if(setsetting != null) setsetting = setsetting.trim();
		if (setsetting.equals("Global")) testsettings(rset);
		rset1.next();
				
		for (int nTcs = 1; nTcs <= NumOfTcs; nTcs++)
		{
			
			if (!(setsetting.equals("Global"))) testsettings(rset1);
		
			Option1 = switchOption1(webdriver);
			
			switch (Option1) 
			{
				case 1:	driver = new InternetExplorerDriver(); break;
				case 2:	driver = new FirefoxDriver(); break;
				case 3:  
					File directory = new File (".");
					chromedriver = directory.getCanonicalPath() + "\\Chromedriver\\chromedriver.exe";
					System.setProperty("webdriver.chrome.driver",chromedriver); 
					service = ChromeDriverService.createDefaultService(); 
					service.start(); driver = new ChromeDriver(service); break;
				default: System.out.println("TestError: Webdriver is not defined!"); break;
			}			

			driver.manage().timeouts().implicitlyWait(synctime, TimeUnit.MILLISECONDS);
			if (driver==null) break; 
			driver.get(baseurl);	
			selenium = new WebDriverBackedSelenium(driver, baseurl);
			//selenium.windowMaximize();
			//selenium.windowFocus();					
		
			TestScript = rset1.getString("testscript");
			TestScript = TestScript.trim();
			System.out.println("StartExecution: " + TestScript + " execution in progress"); 
			if (TestScript.equals("EndOfRow")) break;
			String TestStep = "Select * From [TestWare$] Where [tcid] = '" + TestScript + "';";
			rset1.next();
			rset2 = DBConnection(TestStep);
			int NumOfCols = recurcount;		
			rset2.last();
			int NumOfRows = rset2.getRow();
			System.out.println("Iterations: " + NumOfCols + " iterations of the " + TestScript + " testcase");
			teststatus = "PASSED";
			for (nCols =1; nCols <= NumOfCols; nCols++)
			{
				rset2.first();
				for (int nRows = 1; nRows <= NumOfRows; nRows++)
				{
					Activity = rset2.getString("Activity");
					Activity = Activity.trim();
					if (Activity.equals("EndOfRow")) break;
					Attrname = rset2.getString("Attrname");
					if(Attrname != null)
					{	
						Attrname = Attrname.trim();
						Option2 = switchOption2(Attrname);
					}
					Attrvalue = rset2.getString("Attrvalue");
					if(Attrvalue != null) Attrvalue = Attrvalue.trim();
					Parameter = rset2.getString("Iteration" + nCols);
					//if(Parameter != null) Parameter = Parameter.trim();
					stepstatus = rset2.getString("stepstatus");
					if(stepstatus != null) stepstatus = stepstatus.trim();
					StepDetails = rset2.getString("TestSteps");
					if(StepDetails != null) StepDetails = StepDetails.trim();
					System.out.println("Iteration"+nCols+": "+Activity+":"+Attrname+"="+Attrvalue+":"+Parameter);
					rset2.next();
					
					if (stepstatus.equals("TRUE"))
					{
						ObjectExists = GetElement(false);
						if (ObjectExists==true)
						{
							Option3 = switchOption3(Activity);
						}
						else
						{
							teststatus = "FAILED";
							String fileName = captureScreenshot();
							String UpdateStatus = "Update [TestWare$] Set [screenshot] = '" + fileName + "' Where [testscript] = '" + TestScript + "';";
							DBConnectionUpdate(UpdateStatus);
							break;
						}
					}
					else if (stepstatus.equals("OPTIONAL"))
					{
						ObjectExists = GetElement(true);
						if (ObjectExists==true)
						{
							Option3 = switchOption3(Activity);
						}
						else
						{
							Option3 = 998;
						}
					}
					else 
					{
						Option3 = 999;
					}					
 					
					try 
					{
						switch(Option3) 
						{
							case 1: // type	
								element.sendKeys(Parameter); 
								 break;											
							case 2: // click
								element.click(); Thread.sleep(synctime);
								break;
							case 3: // select
								select = new Select(element); 
								//select.selectByVisibleText(Parameter);
								//select.selectByValue(Parameter); 
								select.selectByIndex(GetIndex(select, Parameter)); 
								break;
							case 4: // clear
								element.clear();
							case 5: // keysEnter		
								element.sendKeys(Keys.ENTER); 
						

							case 10: // keyEventTab 
						        robot = new Robot(); robot.delay(synctime);
								robot.keyPress(KeyEvent.VK_TAB); robot.keyRelease(KeyEvent.VK_TAB);
								break;
							case 11: // keyEventEnter
								robot = new Robot(); robot.delay(synctime);
								robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER); break;						 
							case 12: // alertOk
								switchToAlert(alertTitle); driver.switchTo().alert().accept();
							  	break; 
							case 13: // alertCancel
								switchToAlert(alertTitle); driver.switchTo().alert().dismiss(); 
							  	break; 
							case 14: // alertButton
								switchToAlert(alertTitle); 
								aitxobj.controlClick(alertTitle, driver.switchTo().alert().getText(), Attrvalue);
								break;
								
							case 15: // verifyAlertText
								switchToAlert(alertTitle);
								String alertTextActual = driver.switchTo().alert().getText();
								if (!alertTextActual.equals(Attrvalue))
								{
									String UpdateStatus = "Update [TestWare$] Set [Stepresult" + nCols + "] = '" + alertTextActual + "' Where [TestSteps] = '" + StepDetails + "';";
									DBConnectionUpdate(UpdateStatus);
									String CheckStatus = "Update [TestWare$] Set [Iteration" + nCols + "] = '" + "Failed" + "' Where [TestSteps] = '" + StepDetails + "';";
									DBConnectionUpdate(CheckStatus);	
								}	
								else
								{
									String CheckStatus = "Update [TestWare$] Set [Iteration" + nCols + "] = '" + "Passed" + "' Where [TestSteps] = '" + StepDetails + "';";
									DBConnectionUpdate(CheckStatus);
								}
								break; 
								
							case 21: // selectFrameByName 
								driver.switchTo().defaultContent();
								String[] frames = Attrvalue.split("\\.");
								if (frames.length == 1)
								{
									driver.switchTo().frame(Attrvalue); 
								}
								else if (frames.length == 2)
								{
									driver.switchTo().frame(frames[0]).switchTo().frame(frames[1]);
								}
								Thread.sleep(synctime);
								
								break;
							case 22: // selectFrameByElement
								driver.switchTo().defaultContent();
								driver.switchTo().frame(element); 
								break;					
								
							case 25: // switchWindowByTitle
								try
								{
									selenium.selectWindow(Attrvalue);
								}
								catch (Exception e4) {
									System.out.println("Iteration" + nCols + " : selectWindow:empty=" + Attrvalue + ":null");
									// TODO: handle exception
								}
								for (String handle : driver.getWindowHandles())
						    	{	
									driver.switchTo().window(handle);
						    		String title = driver.switchTo().window(handle).getTitle();
						    		if (title.contains(Attrvalue))
						    		{
						    			driver.switchTo().window(handle); 
						    			break;
						    		}
						    	} 
								break;
							case 26: // closeWindowByTitle
								try
								{
									selenium.selectWindow(Attrvalue);	
								}
								catch (Exception e4) {
									System.out.println("Iteration" + nCols + " : selectWindow:empty=" + Attrvalue + ":null");
									// TODO: handle exception
								}
								for (String handle : driver.getWindowHandles()) 
						    	{
						    		driver.switchTo().window(handle);
						    		String title = driver.switchTo().window(handle).getTitle();
						    		if (title.equals(Attrvalue))
						    		{
						    			driver.switchTo().window(handle).close(); 
						    			break;
						    		}
						    	}
						    	break; 
						    	
							case 27: // closeAllOtherWindows  		
								for (String handle : driver.getWindowHandles()) 
						    	{
						    		driver.switchTo().window(handle);
						    		String title = driver.switchTo().window(handle).getTitle();
						    		if (!(title.equals(Attrvalue))) driver.switchTo().window(handle).close();
						    	}
								break; 
						    	
							case 30: // deleteCookies 
								driver.manage().deleteAllCookies(); 
								break;
								
						    	
							case 31: // closeDriver
								driver.quit(); driver = null;
								break; 
							case 32: // refresh
								driver.navigate().refresh(); 
								break; 
							case 33: // goBack
								driver.navigate().back();  
								driver.navigate().refresh(); break;
								
							case 34: // waitForSync
								Attrvalue = Attrvalue.substring(0, Attrvalue.length()-2);
								Thread.sleep(Integer.parseInt(Attrvalue)); break;
								
							case 998: 
								System.out.println("Optional step failed!");
								break;
								
							case 999: 
								System.out.println("Step execution status is false!");
								break;
								
							default: 
								System.out.println("TestError: Object type is not defined"); 
								break;						
						}						
					}
					catch (Exception e2) {
						System.out.println("TestError: Problem with operation on the Object");
						// TODO: handle exception
					}
				}
			}
			
			String UpdateStatus = "Update [TestWare$] Set [teststatus] = '" + teststatus + "' Where [testscript] = '" + TestScript + "';";
			DBConnectionUpdate(UpdateStatus);
			try
			{
				//driver.quit(); // Close driver at last step				
				//System.out.println("StopExecution: closeDriver:empty=null:null");
				if (webdriver.equals("GC")) service.stop();
			}	
			catch (Exception e3) {
				System.out.println("StopExecution: Web Driver has been closed!");
				// TODO: handle exception
			}
		}
    	stmt.close();	
    	conn.close();
    	saveExcelDocument();
    	jasperReportsGeneration("reports/report1.jasper", "#", "reports/report1.jrprint");
     	closeExcelDocument();
	 }
		
	

	private boolean GetElement(boolean isOptional) throws Exception {
		boolean isExists = true;
		for (int i = 1; i <= retrycount; i++)
		{
			try 
			{
				switch(Option2) 
				{
					case 1: element = driver.findElement(By.name(Attrvalue)); break;
					case 2: element = driver.findElement(By.id(Attrvalue)); break;
					case 3: element = driver.findElement(By.xpath(Attrvalue)); break;
					case 4: element = driver.findElement(By.linkText(Attrvalue)); break;
					case 5: element = driver.findElement(By.cssSelector(Attrvalue)); break;
					case 6:	
						// System.out.println("Warning: The findElement attribute is not used");						
					case -1:
						// System.out.println("Warning: The findElement attribute is not used");
						break;
					default: 
						isExists = false; 
						System.out.println("TestError: The findElement attribute is not defined"); break;
				}
			}
			catch (Exception e1) {
				// System.out.println(e1.getMessage());
				isExists = false;	
				//if(closePopUps()) break;
				String activeTitle = driver.getTitle();
				if (activeTitle.equals("Active and Dismissed Alerts"))
				{
					driver.close();		
					selenium.selectWindow(null);
				}
				if (activeTitle.equals("Completed Report Alert"))
				{
					driver.close();		
					selenium.selectWindow(null);
				}
				
				if (e1.getMessage().startsWith("Unable to find element"))
				{
					System.out.println("TestError: Unable to find element : Retry Count - " + i);
				}
				else if (e1.getMessage().startsWith("No window found"))
				{
					System.out.println("TestError: No window found : Retry Count - " + i);
				}
				else
				{
					System.out.println("TestError: Problem in Object Identification"); 
					break;
				}	
				// TODO: handle exception
			}
			if (isOptional) i = retrycount;	
			if (isExists==true) break;
			Thread.sleep(synctime);			
		}
		return isExists;
	}

	private ResultSet DBConnection(String StrQuery) throws Exception{
		String osarch = System.getProperty("os.arch");
		File directory = new File (".");
		Class.forName( "sun.jdbc.odbc.JdbcOdbcDriver" );
		String Path = directory.getCanonicalPath()+ "\\data\\TestWare.xls";
		try{
			if (osarch.equals("x86"))
			{
				conn = DriverManager.getConnection("jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};DBQ=" + Path + "; readOnly = false","","");
			}
			else if (osarch.equals("amd64"))
			{
				conn = DriverManager.getConnection("jdbc:odbc:Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + Path + "; readOnly = false","","");
			}		
			stmt = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_UPDATABLE);
			rset = stmt.executeQuery(StrQuery);
		}
		catch (Exception e) {
			// TODO: handle exception
			System.out.println(e.getMessage());
		}
		
		return rset;
	}
	
	private void DBConnectionUpdate(String StrQuery) throws ClassNotFoundException, IOException, InterruptedException{
		
		String osarch = System.getProperty("os.arch");
		File directory = new File (".");
		Class.forName( "sun.jdbc.odbc.JdbcOdbcDriver" );
		String Path = directory.getCanonicalPath()+ "\\data\\TestWare.xls";
		try{
			if (osarch.equals("x86"))
			{
				conn = DriverManager.getConnection("jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};DBQ=" + Path + "; readOnly = false","","");
			}
			else if (osarch.equals("amd64"))
			{
				conn = DriverManager.getConnection("jdbc:odbc:Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + Path + "; readOnly = false","","");
			}		

		}
		catch (Exception e) {
			// TODO: handle exception
			System.out.println(e.getMessage());
		}

		try {
			stmt = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_UPDATABLE);
			stmt.executeUpdate(StrQuery);
			conn.setAutoCommit(true);
			saveExcelDocument();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private int switchOption1(String webdriverlcl) throws Exception {
		if (webdriverlcl.equals("IE")) {
			indvar = 1;
		} else if (webdriverlcl.equals("MF")) {
			indvar = 2;
		} else if (webdriverlcl.equals("GC")) {
			indvar = 3;
		} else {
			indvar = 0;
		}
		return indvar;
	}
	
	private int switchOption2(String Attrnamelcl) throws Exception {
		if (Attrnamelcl.equals("name")) {
			indvar = 1;
		} else if (Attrnamelcl.equals("id")) {
			indvar = 2;
		} else if (Attrnamelcl.equals("xpath")) {
			indvar = 3;
		} else if (Attrnamelcl.equals("linkText")) {
			indvar = 4;
		} else if (Attrnamelcl.equals("cssSelector")) {
			indvar = 5;
		} else if (Attrnamelcl.equals("window")) {
			indvar = 6;
		} else if (Attrnamelcl.equals("empty")) {
			indvar = -1;
		} else {
			indvar = 0;
		}
		return indvar;
	}
	
	private int switchOption3(String Activitylcl) throws Exception {
		
		if (Activitylcl.equals("type")) {
			indvar = 1;
		} else if (Activitylcl.equals("click")) {
			indvar = 2;
		} else if (Activitylcl.equals("select")) {
			indvar = 3;
		} else if (Activitylcl.equals("clear")) {
			indvar = 4;
		} else if (Activitylcl.equals("keysEnter")) {
			indvar = 5;
		} else if (Activitylcl.equals("keyEventTab")) {
			indvar = 10;
		} else if (Activitylcl.equals("keyEventEnter")) {
			indvar = 11;
		} else if (Activitylcl.equals("alertOk")) {
			indvar = 12;
		} else if (Activitylcl.equals("alertCancel")) {
			indvar = 13;
		} else if (Activitylcl.equals("alertButton")) {
			indvar = 14;
		} else if (Activitylcl.equals("verifyAlertText")) {
			indvar = 15;
		} else if (Activitylcl.equals("selectFrameByName")) {
			indvar = 21;
		} else if (Activitylcl.equals("selectFrameByElement")) {
			indvar = 22;
		} else if (Activitylcl.equals("switchWindowByTitle")) {
			indvar = 25;
		} else if (Activitylcl.equals("closeWindowByTitle")) {
			indvar = 26;
		} else if (Activitylcl.equals("closeAllOtherWindows")) {
			indvar = 27;
		} else if (Activitylcl.equals("deleteCookies")) {
			indvar = 30;
		} else if (Activitylcl.equals("closeDriver")) {
			indvar = 31; 
		} else if (Activitylcl.equals("refresh")) {
			indvar = 32;
		} else if (Activitylcl.equals("goBack")) {
			indvar = 33;
		} else if (Activitylcl.equals("waitForSync")) {
			indvar = 34;
		} else {
			indvar = 0;
		}
		return indvar;			
	}
	 
	private void testsettings(ResultSet rset) throws Exception {
		baseurl = rset.getString("baseurl");		
		if(baseurl != null) baseurl = baseurl.trim();		
		webdriver = rset.getString("webdriver");
		if(webdriver != null) webdriver = webdriver.trim();
		synctime = rset.getInt("synctime");		
		recurcount = rset.getInt("recurcount");		
		retrycount = rset.getInt("retrycount");				
	}
	
	private int GetIndex(Select selectlcl, String Parameterlcl) throws Exception {
		
		int nItm = 0; String Index = "Index:";
		if (Parameterlcl.startsWith(Index))
		{
			nItm = Integer.parseInt(Parameterlcl.substring(Index.length()));
		}
		else
		{
			List<WebElement> allOptions = selectlcl.getOptions();
			for (WebElement option : allOptions)
			{
				String itemText = option.getText();
				if (itemText.equals(Parameterlcl)) break;
				nItm++;
			}
		}
		return nItm;
	}
	
	private void initializeAutoItX() throws InterruptedException {
		if (osarch.equals("x86"))
		{
			File file = new File("lib/JACOB", "jacob-1.16-M1-x86.dll"); 
	        System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
	        aitxobj = new AutoItX();   
		}
		else if (osarch.equals("amd64"))
			
		{
			File file = new File("lib/JACOB", "jacob-1.16-M1-x64.dll"); 
	        System.setProperty(LibraryLoader.JACOB_DLL_PATH, file.getAbsolutePath());
	        aitxobj = new AutoItX(); 
		}       
 	}
	
	private void switchToAlert(String alertTitle) throws InterruptedException {
		initializeAutoItX();
		if(aitxobj.winExists(alertTitle))
		{
			aitxobj.winActivate(alertTitle); 			
		}
 	}
	
	private String captureScreenshot() {
		String fileName = TestScript + " " + getDateAndHour() + ".png";	
		try {
		File imgFile = new File("images/" + fileName);
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);	
		FileUtils.copyFile(scrFile, imgFile);
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("captureScreenshot");
		return fileName;
	}
	
	// Get current Date, Hour, Minute, Second, Millisecond; 
	// This method is used to create the screen shot file name 
	private String getDateAndHour() { 
		String today; 
		DateFormat dateFormat = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss"); 
		Calendar calendar = Calendar.getInstance(); 
		today = dateFormat.format(calendar.getTime()); 
		return today; 
	} 
	
	private void jasperReportsGeneration(String jasperFile, String HEADER_LINK_VAL, String jrprintFile) throws JRException, IOException
	{
		long start = System.currentTimeMillis();
		File directory = new File (".");
		//Preparing parameters 
		String REPORT_DIR_VAL = directory.getCanonicalPath();
		Map parameters = new HashMap();
		parameters.put("HEADER_LINK", HEADER_LINK_VAL);
		parameters.put("REPORT_DIR", REPORT_DIR_VAL);
		try{
			JasperFillManager.fillReportToFile(jasperFile, parameters, getDataSource());
		}
		catch (JRException e) {
			// TODO: handle exception
			System.out.println("Exception");
		}
		
		System.err.println("Filling time : " + (System.currentTimeMillis() - start));
		JasperExportManager.exportReportToHtmlFile(jrprintFile);
		System.err.println("HTML creation time : " + (System.currentTimeMillis() - start));
		System.out.println("End");
	}
	
	private JRXlsDataSource getDataSource() throws JRException, IOException
	  {	
		File directory = new File (".");
		//Preparing parameters 
		String TestWare = directory.getCanonicalPath()+ "\\data\\TestWare.xls";
		JRXlsDataSource ds;
	    try
	    {
		  ds = new JRXlsDataSource(JRLoader.getLocationInputStream(TestWare));
	      ds.setUseFirstRowAsHeader(true);
	    }
	    catch (IOException e)
	    {
	      throw new JRException(e);
	    }
	    return ds;
	 }
	
	private void openExcelDocument() throws IOException, InterruptedException {
		  File file = File.createTempFile("openExl",".vbs");
	      file.deleteOnExit();
	      FileWriter fw = new java.io.FileWriter(file);
	      File directory = new File (".");
	      String filePath = directory.getCanonicalPath()+ "\\data\\TestWare.xls"; 
	      String vbsXls = 
	    	  	"On Error Resume Next\n" +
	    	  	"FilePath = " + "\"" + filePath + "\"\n" +
	    	  	"Set ExcelApp = CreateObject(\"Excel.Application\")\n" +
	    	  	"Set ExcelApp = GetObject(, \"Excel.Application\")\n" +
	      		"ExcelApp.Application.Visible = True\n" +
	      		"Set ExcelBook = ExcelApp.Workbooks.Open(FilePath)\n" +
	      		"Set ExcelBook = Nothing\n" +
	      		"Set ExcelApp = Nothing";
	      fw.write(vbsXls);
	      fw.close();
	      System.out.println(file.getPath());
	      Runtime.getRuntime().exec("cscript //NoLogo " + file.getPath());	
	      Thread.sleep(synctime);
	}
	
	private void closeExcelDocument() throws IOException, InterruptedException {
		  File file = File.createTempFile("closeExl",".vbs");
	      file.deleteOnExit();
	      FileWriter fw = new java.io.FileWriter(file);
	      File directory = new File (".");
	      String filePath = directory.getCanonicalPath()+ "\\data\\TestWare.xls"; 
	      String vbsXls = 
	    	  	"On Error Resume Next\n" +
	    	  	"Set ExcelApp = GetObject(, \"Excel.Application\")\n" +
	      		"If (Err.Number = 0) Then\n" +
	      		"ExcelApp.Quit\n" +
	      		"End if\n" +
	      		"Set ExcelApp = Nothing";
	      fw.write(vbsXls);
	      fw.close();
	      System.out.println(file.getPath());
	      Runtime.getRuntime().exec("cscript //NoLogo " + file.getPath());	
	      Thread.sleep(synctime);
	}
	
	private void saveExcelDocument() throws IOException, InterruptedException {
		  File file = File.createTempFile("saveExl",".vbs");
	      file.deleteOnExit();
	      FileWriter fw = new java.io.FileWriter(file);
	      File directory = new File (".");
	      String filePath = directory.getCanonicalPath()+ "\\data\\TestWare.xls"; 
	      String vbsXls = 
	    	  	"On Error Resume Next\n" +
	    	  	"Set ExcelWorkBook = GetObject(, \"Excel.Application\")\n" +
	      		"If (Err.Number = 0) Then\n" +
	      		"For Each aWorkbook In ExcelWorkBook.Workbooks\n" +
	      		"If aWorkbook.Saved = False Then\n" +
	      		"aWorkbook.Save\n" +
	      		"End If\n" +
	      		"Next\n" +
	      		"Set ExcelWorkBook = Nothing\n" +
	      		"End If";	      		
	      fw.write(vbsXls);
	      fw.close();
	      System.out.println(file.getPath());
	      Runtime.getRuntime().exec("cscript //NoLogo " + file.getPath());
	      Thread.sleep(synctime);
	}
	
	private boolean verifyElementProperty(WebElement eleObject, String compareProperty, String expectedValue) {
		String actualValue = eleObject.getAttribute(compareProperty);
		if (expectedValue.equals(actualValue))
		{
			System.out.println("Checkpoint Passed");
			return true;
		}
		else
		{
			System.out.println("Checkpoint Failed");
			return false;
		}			
	}
}



