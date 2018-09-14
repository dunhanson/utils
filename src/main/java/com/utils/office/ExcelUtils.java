package com.utils.office;


import jdk.internal.util.xml.impl.Input;
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
	    return readWorkbookByLimit(in, 0, 0, 0);
	}

    /**
     * 读取workbook
     * @param in InputStream
     * @param sheetNum sheet序号
     * @return
     */
	public static List<Object[]> readWorkbookBySheetNum(InputStream in, int sheetNum) {
		return readWorkbookByLimit(in, sheetNum, 0, 0);
	}

	/**
	 * 读取workbook
	 * @param in InputStream
	 * @param sheetName sheet名称
	 * @return
	 */
	public static List<Object[]> readWorkbookBySheetName(InputStream in, String sheetName) {
		return readWorkbookByLimit(in, sheetName, 0, 0);
	}

	/**
	 * 读取workbook
	 * @param in InputStream
	 * @param startNum 开始行号
	 * @param endNum 结束行号
	 * @return
	 */
	public static List<Object[]> readWorkbookByLimit(InputStream in, int startNum, int endNum) {
		return readWorkbookByLimit(in, 0, startNum, endNum);
	}

    /**
     * 读取workbook
     * @param in InputStream
     * @param sheetName sheet名称
     * @param startNum 开始行号
     * @param endNum 结束行号
     * @return
     */
    public static List<Object[]> readWorkbookByLimit(InputStream in, String sheetName, int startNum, int endNum) {
        return readWorkbookByLimit(in, getSheetNum(in, sheetName), startNum, endNum);
    }

    /**
     * 读取workbook
     * @param in InputStream
     * @param sheetNum sheet名称
     * @param startNum 开始行号
     * @param endNum 结束行号
     * @return
     */
    public static List<Object[]> readWorkbookByLimit(InputStream in, int sheetNum, int startNum, int endNum) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(in);
            return readSheetByLimit(wb.getSheetAt(sheetNum), startNum, endNum);
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
		return readWorkbookByLimit(filePath, 0, 0, 0);
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param sheetNum sheet名称
	 * @return
	 */
	public static List<Object[]> readWorkbookBySheetNum(String filePath, int sheetNum) {
		return readWorkbookByLimit(filePath, sheetNum, 0, 0);
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param sheetName sheet名称
	 * @return
	 */
	public static List<Object[]> readWorkbookBySheetName(String filePath, String sheetName) {
		return readWorkbookByLimit(filePath, sheetName, 0, 0);
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param startNum 开始行号
	 * @param endNum 结束行号
	 * @return
	 */
	public static List<Object[]> readWorkbookByLimit(String filePath, int startNum, int endNum) {
		return readWorkbookByLimit(filePath, 0, startNum, endNum);
	}

	/**
	 * 读取workbook
	 * @param filePath 文件路径
	 * @param sheetName sheet名称
	 * @param startNum 开始行号
	 * @param endNum 结束行号
	 * @return
	 */
	public static List<Object[]> readWorkbookByLimit(String filePath, String sheetName, int startNum, int endNum) {
		return readWorkbookByLimit(filePath, getSheetNum(filePath, sheetName), startNum, endNum);
	}

    /**
     * 读取workbook
     * @param filePath 文件路径
     * @param sheetNum sheet序号
     * @param startNum 开始行号
     * @param endNum 结束行号
     * @return
     */
    public static List<Object[]> readWorkbookByLimit(String filePath, int sheetNum, int startNum, int endNum) {
        try(InputStream in = new FileInputStream(filePath)) {
            return readWorkbookByLimit(in, sheetNum, startNum, endNum);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new ArrayList<>();
    }

	/**
	 * 读取Sheet数据
	 * @param sheet Sheet对象
	 * @return
	 */
	public static List<Object[]> readSheet(Sheet sheet) {
		return readSheetByLimit(sheet, 0, sheet.getPhysicalNumberOfRows());
	}

	/**
	 * 读取Sheet数据
	 * @param sheet Sheet
	 * @param startNum 开始行号
	 * @param endNum 结束行号
	 * @return
	 */
	public static List<Object[]> readSheetByLimit(Sheet sheet, int startNum, int endNum) {
		List<Object[]> list = new ArrayList<>();
        //计算区间
		if(startNum == 0) startNum = 1;
		if(endNum == 0) endNum = sheet.getPhysicalNumberOfRows();
		//遍历区间
		for(int i = startNum - 1; i < endNum; i++) {
			Row row = sheet.getRow(i);
			if(row == null) {
			    continue;
            }
			int cellNums = row.getPhysicalNumberOfCells();
			Object[] arr = new Object[cellNums];
			for(int j = 0; j < cellNums; j++) {
				arr[j] = row.getCell(j);
			}
			list.add(arr);
		}
		return list;
	}

    /**
     * 获取Sheet序号
     * @param filePath 文件路径
     * @param sheetName Sheet名称
     * @return
     */
	public static int getSheetNum(String filePath, String sheetName) {
        try(InputStream in = new FileInputStream(filePath)) {
            Workbook wb = WorkbookFactory.create(new FileInputStream(filePath));
            return wb.getSheetIndex(sheetName);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return -1;
    }


    /**
     * 获取Sheet序号
     * @param in InputStream
     * @param sheetName Sheet名称
     * @return
     */
    public static int getSheetNum(InputStream in, String sheetName) {
        try {
            return WorkbookFactory.create(in).getSheetIndex(sheetName);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return -1;
    }

    /**
     * 获取Sheet名称
     * @param filePath 文件路径
     * @param sheetNum Sheet序号
     * @return
     */
    public static String getSheetName(String filePath, int sheetNum) {
        try(InputStream in = new FileInputStream(filePath)) {
            Workbook wb = WorkbookFactory.create(in);
            return wb.getSheetName(sheetNum);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取Sheet名称
     * @param in InputStream
     * @param sheetName Sheet名称
     * @return
     */
    public static String getSheetName(InputStream in, int sheetName) {
        try {
            return WorkbookFactory.create(in).getSheetName(sheetName);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

}
