package TestDataprovider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProvider_apache_02 {
	WebDriver driver;
	// declaration for xlsx workbook
	private XSSFSheet sh;
	private XSSFWorkbook wb;
	private XSSFCell Cell;
	private XSSFRow Row;

	@BeforeClass
	public void Initialisation() throws Exception {
		// driver = new FirefoxDriver();
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\aravind\\Desktop\\Selenium_Class\\LatestBrowserJar\\Chrome_2.40\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("http://www.store.demoqa.com");
		Thread.sleep(3000);
		driver.manage().window().maximize();
	}

	@AfterClass
	public void endoperation() {
		driver.quit();
	}

	@DataProvider(name = "Authentication")
	public Object[][] Authentication() throws Exception {
		Object[][] testObjArray = getTableArray(
				"C:\\Users\\aravind\\Selenium_HITS_Sessions\\Maven_Dataprovider_TestNG\\src\\test\\resources\\Data\\TestData.xlsx",
				"Sheet1");
		return (testObjArray);
	}

	@DataProvider(name = "DataSet")
	public Object[][] DataForForm() throws Exception {
		Object[][] testObjArray = getTableArray(
				"C:\\Users\\aravind\\Selenium_HITS_Sessions\\Maven_Dataprovider_TestNG\\src\\test\\resources\\Data\\TestData.xlsx",
				"Sheet1");
		return (testObjArray);
	}

	@Test(dataProvider = "Authentication")
	public void Registration_data(String sUserName, String sPassword) throws Exception {
		driver.findElement(By.xpath(".//*[@id='account']/a")).click();
		driver.findElement(By.id("log")).sendKeys(sUserName);
		Thread.sleep(3000);
		System.out.println(sUserName);
		driver.findElement(By.id("pwd")).sendKeys(sPassword);
		Thread.sleep(3000);
		System.out.println(sPassword);
		driver.findElement(By.id("login")).click();
		Thread.sleep(3000);
		System.out.println(" Login Successfully, now it is the time to Log Off buddy.");
	}

	// Read the Excel data
	public Object[][] getTableArray(String FilePath, String SheetName) throws Exception {
		String[][] tabArray = null;
		try {
			FileInputStream ExcelFile = new FileInputStream(FilePath);
			// Access the required test data sheet
			wb = new XSSFWorkbook(ExcelFile);
			sh = wb.getSheet(SheetName);
			// Optional Temporary Variable which is going to store the cell
			// values
			int ci, cj;
			int totalRows = sh.getLastRowNum();
			int noOfColumns = sh.getRow(totalRows).getLastCellNum();
			int col = noOfColumns - 1;
			System.out.println("The total cells(column) : " + col);
			tabArray = new String[totalRows][col];
			ci = 0;// Row values
			for (int i = 1; i <= totalRows; i++, ci++) {
				cj = 0;
				// Column values
				for (int j = 1; j <= col; j++, cj++) {
					// data from the excel
					tabArray[ci][cj] = getCellData(i, j);
					System.out.println("The values for i and j : " + tabArray[ci][cj]);
				}
			}
		} catch (FileNotFoundException e) {
			System.out.println("Could not read the Excel sheet");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("Could not read the Excel sheet");
			e.printStackTrace();
		}
		return (tabArray);
	}

	public String getCellData(int RowNum, int ColNum) throws Exception {
		try {
			Cell = sh.getRow(RowNum).getCell(ColNum);
			System.out.println("print the cell : " + Cell);
			/*
			 * int dataType = Cell.getCellType(); if (dataType == 3) { return
			 * ""; }else{
			 */
			String CellData = Cell.getStringCellValue();
			System.out.println("The cell data value : " + CellData);
			return CellData;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			throw (e);

		}
	}
}
