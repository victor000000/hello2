import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 资金状况数据
 * @author vic
 *
 */
public class ZJZK {

	protected static HSSFWorkbook workbook = null;

	/**
	 * 得到上日结存
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getSRJC(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 11, 2);
	}
	
	/**
	 * 得到上日存取合计
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getDRCQHJ(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 12, 2);
	}
	
	/**
	 * 得到当日盈亏
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getDRYK(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 13, 2);
	}
	
	/**
	 * 当日手续费
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getDRSXF(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 14, 2);
	}
	
	/**
	 * 当日结存
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getDRJC(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 15, 2);
	}
	
	/**
	 * 客户权益
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getKHQY(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 11, 7);
	}
	
	/**
	 * 质押金
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getZYJ(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 12, 7);
	}
	
	/**
	 * 保证金占用
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getBZJZY(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 13, 7);
	}
	
	/**
	 * 可用资金
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getKYZJ(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 14, 7);
	}
	
	
	/**
	 * 风险度
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getFXD(String filePath) throws IOException  {
		initWorkbook(filePath);
		String s = getString(0, 15, 7);
		double d = Double.valueOf(s.substring(0, s.length() - 1 - 1));
		d = d / 100.0;
		return d;
	}
	
	/**
	 * 追加保证金
	 * @param filePath
	 * @return
	 * @throws IOException
	 */
	public static double getZJBZJ(String filePath) throws IOException  {
		initWorkbook(filePath);
		return getDouble(0, 16, 7);
	}
	
	protected static void initWorkbook(String filePath) throws IOException {
		FileInputStream file = new FileInputStream(new File(filePath));
		workbook = new HSSFWorkbook(file);
	}
	
	protected static double getDouble(int sheet, int row, int col) {
		return workbook.getSheetAt(sheet).getRow(row).getCell(col).getNumericCellValue();
	}
	
	protected static String getString(int sheet, int row, int col) {
		return workbook.getSheetAt(sheet).getRow(row).getCell(col).getStringCellValue();
	}
	
}
