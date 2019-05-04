package step.com.test.javaexamples;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class JavaExample12 {
	
	//This example talks about excel file operations
	
	public Workbook fileOpen(String Path, String FileName){
		Workbook resultFileWorkBook = null;
		InputStream inp = null;
		
		try{
			inp = new FileInputStream(Path + FileName);
			resultFileWorkBook = WorkbookFactory.create(inp);
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
		
		return resultFileWorkBook;
	}
	
	public void fileSaveClose(Workbook WB, String Path, String FinalResultFileName) {
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(Path + FinalResultFileName);
			WB.write(fileOut);
			fileOut.close(); 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void writeData(Sheet SH, int RowNum, int CellNum, String Data) {
		Row row = null;     
		Cell cell = null;
		row = SH.getRow(RowNum);
		if(row == null){
			row = SH.createRow(RowNum);
		}
		cell = row.getCell(CellNum);
		if(cell == null){
			cell = row.createCell(CellNum);
		}
		cell.setCellValue(Data);
	}
	
	public void compareResults(Sheet sheet, int resRow, int expCell, int actCell, int resCell) {
		Row compareResultRow = sheet.getRow(resRow);
		Cell compareResultExpected = compareResultRow.getCell(expCell);
		Cell compareResultActual = compareResultRow.getCell(actCell);
		String expectedValue = compareResultExpected.getStringCellValue();
		String actualValue = compareResultActual.getStringCellValue();
		if(expectedValue.contentEquals(actualValue)){
			writeData(sheet, resRow, resCell, "Pass");
		}else{
			writeData(sheet, resRow, resCell, "Fail");
		}
	}

	public static void main(String args[]){
		JavaExample12 example = new JavaExample12();
		Workbook fileOne = example.fileOpen("C:\\Working\\StepIn2It\\StepIn2ItInputFiles\\", "ExcelFileExample.xlsx");
		Sheet s1 = fileOne.getSheetAt(0);
		example.writeData(s1, 1, 0, "Test Case 1");
		example.writeData(s1, 1, 2, "successss");
		example.compareResults(s1, 1, 1, 2, 3);
		example.fileSaveClose(fileOne, "C:\\Working\\StepIn2It\\StepIn2ItInputFiles\\", "ExcelFileExample.xlsx");
	}
}