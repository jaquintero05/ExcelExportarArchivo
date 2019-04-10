package excelexportararchivo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

	
public class Excel {
	
	private By BarraBusqueda = By.name("q");
	private By BotonBuscar = By.name("btnK");
	private WebDriver driver;
	
	
	public Excel(WebDriver driver) {
		
	this.driver = driver;
	}
	
	public  void buscarExcel(){
		        
		    	try { 	
		    	    String rutaArchivoExcel = "C:\\Selenium\\practica.xlsx";
		            FileInputStream inputStream = new FileInputStream(new File(rutaArchivoExcel));
		            Workbook workbook = new XSSFWorkbook(inputStream);
		            Sheet firstSheet = workbook.getSheetAt(0);
		            Iterator iterator = firstSheet.iterator();
		            
		            DataFormatter formatter = new DataFormatter();// captura los datos de cada celda
		            while (iterator.hasNext()) {
		                Row nextRow = (Row) iterator.next(); //fila row
		                Iterator cellIterator = nextRow.cellIterator();
		                
		                
		                while(cellIterator.hasNext()) {
		                    Cell cell = (Cell) cellIterator.next();
		                    String contenidoCelda = formatter.formatCellValue(cell);
		                    System.out.println("celda: " + contenidoCelda);
		                    driver.get("http:\\www.google.com.uy");
	                 		 driver.findElement(BarraBusqueda).sendKeys(contenidoCelda);;
	                 		 driver.findElement(BotonBuscar).submit();
	                 		 WebElement element = (new WebDriverWait(driver, 10))
	              				  .until(ExpectedConditions.presenceOfElementLocated((By.id("search"))));
		                
	     
	                 		try {
	                 			if (ExpectedConditions.invisibilityOf(driver.findElement(By.xpath("//h1[contains(text(),'Resultados de b�squeda')]"))) != null) {
	                 				nextRow.createCell(1).setCellValue("Si Encontrado");
	                 			}	                 			
	                 		 }catch (Exception e) {
	                 			nextRow.createCell(1).setCellValue("No encontrado");
	                 		 }
		                }
		                
		            
		                }
		           
		            inputStream.close();
		     		FileOutputStream outputStream =new FileOutputStream(new File("C:\\Selenium\\practica.xlsx"));
		     		workbook.write(outputStream);
			        outputStream.close();
			        System.out.println("Done");
			        driver.close();
		           }
		           catch (IOException e) {
		        	   
		           }
		    	
		    	
		    	
		        } 
		    
		    }



