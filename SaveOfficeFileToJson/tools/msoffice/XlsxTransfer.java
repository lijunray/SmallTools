package tools.msoffice.saveofficefiletojson;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.Gson;

public class XlsxTransfer {
	
	private static ArrayList<HashMap<String, String>> readFile(String filePath, HashMap<String, Integer> columnMap, int sheetNumber) throws IOException {
	    
		FileInputStream fileInputStream = new FileInputStream(new File(filePath));
	    ArrayList<HashMap<String, String>> array = new ArrayList<>();
	    XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream);
	    XSSFSheet sheet = xssfWorkbook.getSheetAt(sheetNumber - 1);
	    if (columnMap == null) {
	    	ArrayList<String> keys = new ArrayList<>();
	    	for (int i = 0; i <= sheet.getLastRowNum(); i++) {
	    		XSSFRow row = sheet.getRow(i);
	    		HashMap<String, String> map = new HashMap<>();
	    		if (i == 0) {
	    			Iterator<Cell> cellIterator = row.cellIterator();
	    			while (cellIterator.hasNext()) {
	    				keys.add(checkNull(cellIterator.next()));
	    			}
	    		}
	    		else {
	    			for (int j = 0; j < row.getLastCellNum(); j++) {
	    				map.put(keys.get(j), checkNull(row.getCell(j)));
	    			}
	    			array.add(map);
	    		}
	    	}
	    }
	    else {
	    	for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
	    		XSSFRow row = sheet.getRow(i);
	    		HashMap<String, String> map = new HashMap<>();
	    		Iterator<String> keysIterator = columnMap.keySet().iterator();
	    		while (keysIterator.hasNext()) {
	    			String key = keysIterator.next();
	    			map.put(key, checkNull(row.getCell(columnMap.get(key) - 1)));
	    		}
	    		array.add(map);
	    	}
	    }
	    return array;
	    	
	}
	
	private static String checkNull (Cell cell) {
		if (cell == null) {
			return "";
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		return cell.getStringCellValue();
	}
	
	public static boolean saveAsJsonFile (ArrayList<HashMap<String, String>> array, String fileName) throws IOException {
		File file = new File(fileName);
		if (file.exists()) {
			throw new IOException("File exists!");
		}
		Gson gson = new Gson();
		String json = gson.toJson(array);
		BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), "UTF-8"));
		writer.write(json);
		writer.close();
		return true;
	}
}
