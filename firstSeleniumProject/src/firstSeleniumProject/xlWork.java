package firstSeleniumProject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Scanner;


public class xlWork {

	public static void main(String[] args) {
		//System.out.println("Hello world");
		
		System.out.println("Enter Product Name: ");
		Scanner scanner = new Scanner(System.in);
		String user_input = scanner.next();
		
		System.setProperty("webdriver.chrome.driver", "C:/Users/islam/Desktop/selenium files/chromedriver-win64/chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		
        //driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(15));

		driver.get("https://zap.co.il");
		
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        
        WebElement searchInput = wait.until(ExpectedConditions.visibilityOfElementLocated((By.id("acSearch-input"))));
        
        searchInput.sendKeys(user_input);
        
        WebElement searchButton =  wait.until(ExpectedConditions.visibilityOfElementLocated((By.id("acSubmitSearch"))));
        searchButton.click();
        
        List<WebElement> list_of_names = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.className("ModelTitle")));
//        WebElement temp_name = list_of_names.get(0);
//        System.out.println("Here is the temp-name print : "+ temp_name.getText());
        
        List<WebElement> list_of_prices = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.className("ModelDetailsContainer")));
//        WebElement temp_price = list_of_prices.get(0);
        
//        System.out.println("Here is the temp-price print : "+ temp_price.getText());
	

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Products");

        Row headerRow = sheet.createRow(0);
        Cell headerCell1 = headerRow.createCell(0);
        headerCell1.setCellValue("Type");
        
        Cell headerCell2 = headerRow.createCell(1);
        headerCell2.setCellValue("Price");

        for(int i = 0 ; i < list_of_names.size() ; i++) {
        	Row row = sheet.createRow(i+1);
            row.createCell(0).setCellValue(list_of_names.get(i).getText());
            row.createCell(1).setCellValue(list_of_prices.get(i).getText());
        }
        
        // Resize columns to fit content
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("products.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Closing the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Excel file created successfully!");
	}

}
