package com.utils.office;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel操作工具类
 * @author dunhanson
 * @since 2018-09-13 21:17:00
 *
 */
public class ExcelUtils {

	/**
	 * 读取workbook
	 * @param in InputStream
	 * @return
	 */
	public static List<Object[]> readWorkbook(InputStream in) {
		return readWorkbook(in, 0);
	}

	public static List<Object[]> readWorkbook(InputStream in, int sheetNum) {
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(in);
			Sheet sheet = wb.getSheetAt(sheetNum);
			return readSheet(sheet, 0, sheet.getPhysicalNumberOfRows());
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ArrayList<>();
	}

	/**
	 * 读取workbook
	 * @param in InputStream
	 * @param sheetName sheet名称
	 * @return
	 */
	public static List<Object[]> readWorkbook(InputStream in, String sheetName) {
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(in);
			Sheet sheet = wb.getSheet(sheetName);
			return readSheet(sheet, 0, sheet.getPhysicalNumberOfRows());
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ArrayList<>();
	}

	/**
	 * 读取workbook
	 * @param in InputStream
	 * @param sheetName sheet名称
	 * @param startRowNum 开始行号
	 * @param endRowNum 结束行号
	 * @return
	 */
	public static List<Object[]> readWorkbook(InputStream in, String sheetName, int startRowNum, int endRowNum) {
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(in);
			return readSheet(wb.getSheet(sheetName), startRowNum, endRowNum);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ArrayList<>();
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @return
	 */
	public static List<Object[]> readWorkbook(String filePath) {
		return readWorkbook(filePath, 0);
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param sheetNum sheet名称
	 * @return
	 */
	public static List<Object[]> readWorkbook(String filePath, int sheetNum) {
		try(InputStream in = new FileInputStream(filePath)) {
			Workbook wb = WorkbookFactory.create(in);
			return readSheet(wb.getSheetAt(sheetNum));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ArrayList<>();
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param sheetName sheet名称
	 * @return
	 */
	public static List<Object[]> readWorkbook(String filePath, String sheetName) {
		try(InputStream in = new FileInputStream(filePath)) {
			Workbook wb = WorkbookFactory.create(in);
			return readSheet(wb.getSheet(sheetName));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ArrayList<>();
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param sheetName sheet名称
	 * @param startRowNum 开始行号
	 * @param endRowNum 结束行号
	 * @return
	 */
	public static List<Object[]> readWorkbook(String filePath, String sheetName, int startRowNum, int endRowNum) {
		try(InputStream in = new FileInputStream(filePath)) {
			Workbook wb = WorkbookFactory.create(in);
			return readSheet(wb.getSheet(sheetName), startRowNum, endRowNum);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new ArrayList<>();
	}

	/**
	 * 读取Sheet数据
	 * @param sheet
	 * @return
	 */
	public static List<Object[]> readSheet(Sheet sheet) {
		return readSheet(sheet, 0, sheet.getPhysicalNumberOfRows());
	}

	/**
	 * 读取Sheet数据
	 * @param sheet Sheet
	 * @param startRowNum 开始行号
	 * @param endRowNum 结束行号
	 * @return
	 */
	public static List<Object[]> readSheet(Sheet sheet, int startRowNum, int endRowNum) {
		List<Object[]> list = new ArrayList<>();
		//总行数
		int rowRums= sheet.getPhysicalNumberOfRows();
		for(int i = startRowNum - 1; i < endRowNum; i++) {
			Row row = sheet.getRow(i);
			if(row == null) {
			    continue;
            }
			//总列数
			int cellNums = row.getPhysicalNumberOfCells();
			Object[] arr = new Object[cellNums];
			for(int j = 0; j < cellNums; j++) {
				arr[j] = row.getCell(j);
			}
			list.add(arr);
		}
		return list;
	}

}
