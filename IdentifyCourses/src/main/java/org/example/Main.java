package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;


public class Main {
    public static void main(String[] args) throws IOException {

        WebDriver driver = new ChromeDriver();
        driver.get("https://www.coursera.org");
        driver.manage().window().maximize();

        WebDriverWait wait=new WebDriverWait(driver, Duration.ofSeconds(10));

        //Locating the SearchBar
        driver.findElement(By.xpath("//input[@id='search-autocomplete-input']")).sendKeys("Web Development courses");
        Actions act=new Actions(driver);
        act.keyDown(Keys.ENTER).keyUp(Keys.ENTER).perform();

        //Applying Filters
        WebElement level=   wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[text()='Level']")));
        level.click();
        WebElement beginner=  wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='Beginner']")));
        beginner.click();
        WebElement view1= wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@class='cds-149 cds-button-disableElevation cds-button-primary css-1fsvlah']")));
        view1.click();
        WebElement language= wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[text()='Language']")));
        language.click();
        WebElement english=wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='English']")));
        english.click();
        WebElement view2= wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[@class='cds-button-label' and text()='View']")));
        view2.click();


        List<WebElement> titles = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//h3[@class='cds-CommonCard-title css-6ecy9b']")));
        System.out.println(titles.get(0).getText());
        List<WebElement> ratings = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//span[@class='css-4s48ix']")));
        System.out.println("Rating- "+ratings.get(0).getText());
        List<WebElement> duration = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//p[@class='css-vac8rf']/parent::div[@class='cds-CommonCard-metadata']")));
        System.out.println(duration.get(0).getText().replaceAll(".*?(\\d.*)", "$1"));
        System.out.println(titles.get(1).getText());
        System.out.println("Rating- "+ratings.get(1).getText());
        System.out.println(duration.get(1).getText().replaceAll(".*?(\\d.*)", "$1"));

        //Clearing the Filters
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='Clear all']"))).click();

        level.click();
        List<WebElement> levels = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//span[@class='cds-checkboxAndRadio-labelContent css-tvqrra']")));

        List<WebElement> levelscount = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//span[@class='cds-checkboxAndRadio-labelContent css-tvqrra']//span[@class='css-s63saa']")));


        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Course Data");

        int rowNum = 0;

// Header for Levels
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Level Names");
        headerRow.createCell(1).setCellValue("Count");

// Write Levels data in Excel
        for (int i = 0; i < levels.size(); i++) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(levels.get(i).getText().replaceAll("\\(.*", ""));
            row.createCell(1).setCellValue(levelscount.get(i).getText());
        }
        //Closing the levels dropdown
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View']"))).click();

// Add a blank row for separation
        rowNum++;

// LANGUAGES
        language.click();
        List<WebElement> languages = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//span[@class='cds-checkboxAndRadio-labelContent css-tvqrra']/span")));

        List<WebElement> languagescount = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//span[@class='cds-checkboxAndRadio-labelContent css-tvqrra']//span[@class='css-s63saa']")));


// Header for Languages
        Row langHeaderRow = sheet.createRow(rowNum++);
        langHeaderRow.createCell(0).setCellValue("Language Names");
        langHeaderRow.createCell(1).setCellValue("Count");

// Write Languages data in Excel
        for (int i = 0; i < languages.size(); i++) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(languages.get(i).getText().replaceAll("\\(.*", ""));
            row.createCell(1).setCellValue(languagescount.get(i).getText());
        }

// Auto-size columns
        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

// Save to file
        try (FileOutputStream fileOut = new FileOutputStream("CourseEraData.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();

        System.out.println("Excel file created successfully with Levels and Languages!");
        //Closing the language dropdown
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='View']"))).click();

        driver.navigate().to("https://www.coursera.org");
        //Scrolling down to Footer
        Actions action = new Actions(driver);
        WebElement enterprise = driver.findElement(By.xpath("//a[text()='For Enterprise']"));
        action.sendKeys(Keys.END).perform();

        //Clicking For Enterprise
        enterprise.click();

        //Navigating to For Universities
        driver.findElement(By.xpath("//a[@href='https://www.coursera.org/campus']")).click();

        //Filling the form
        driver.findElement(By.xpath("//input[@id='FirstName']")).sendKeys("John");
        driver.findElement(By.xpath("//input[@id='LastName']")).sendKeys("Smith");
        driver.findElement(By.xpath("//input[@id='Email']")).sendKeys("johnsmithgmail.com");
        driver.findElement(By.xpath("//input[@id='Phone']")).sendKeys("9876543210");

        WebElement d1 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@id='Institution_Type__c']")));
        d1.click();
        Select s1 = new Select(d1);
        s1.selectByValue("University/4 Year College");
        driver.findElement(By.xpath("//input[@id='Company']")).sendKeys("JNTU");

        WebElement d2 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@id='Title']")));
        d2.click();
        Select s2 = new Select(d2);
        s2.selectByValue("Professor");

        WebElement d3 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@id='Department']")));
        d3.click();
        Select s3 = new Select(d3);
        s3.selectByValue("Student Affairs");

        WebElement d4 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@id='Self_Reported_Needs__c']")));
        d4.click();
        Select s4 = new Select(d4);
        s4.selectByValue("Courses for myself");

        WebElement d5 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@id='Country']")));
        d5.click();
        Select s5 = new Select(d5);
        s5.selectByValue("India");

        WebElement d6 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//select[@id='State']")));
        d6.click();
        Select s6 = new Select(d6);
        s6.selectByVisibleText("Andhra Pradesh");

        driver.findElement(By.xpath("//button[@type='submit']")).click();

        //Capturing the error message
        String errormsg = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='ValidMsgEmail']"))).getText();
        System.out.println(errormsg);

        TakesScreenshot ts=(TakesScreenshot)driver;
        File tempimg=ts.getScreenshotAs(OutputType.FILE);
        File img=new File(System.getProperty("user.dir")+"\\screenshot\\coursera.png");
        tempimg.renameTo(img);

        driver.navigate().to("https://www.coursera.org");
        driver.quit();


    }

}
