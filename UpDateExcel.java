
package jlfx;
 
import jxl.*;
import jxl.format.UnderlineStyle;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.Boolean;
 
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
 
import java.io.*;
 
public class UpdateExcel {
 
	/**
	 * @param exlFile
	 * @param sheetIndex
	 * @param col
	 * @param row
	 * @param value
	 * @throws Exception
	 */
	public static void updateExcel(File exlFile, int sheetIndex, int col,
			int row, Double value) throws Exception {
		FileInputStream fis = new FileInputStream(exlFile);
		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		// workbook.
		HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
 
		HSSFRow r = sheet.getRow(row);
		HSSFCell cell = r.createCell(col);
		
		cell.setCellValue(value);
		fis.close();
		FileOutputStream fos = new FileOutputStream(exlFile);
		workbook.write(fos);
		fos.close();
	}
 
}
