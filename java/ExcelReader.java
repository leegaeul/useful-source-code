import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * <h4>ExcelReader.java</h4>
 *
 * @author <a href="mailto:gaeul.lee@kt.com"><b>이가을</b></a>
 * @since 2017. 12. 26.
 * @note:
 * <pre>
    다음과 같이 사용할 수 있다.
   ExcelReader reader = ExcelReader.getInstance();
	 try {
	  reader.readXls(new File("resources/jxlrwtest.xls"));
	  reader.readXlsx(new File("resources/receivers.xlsx"));
	 } catch (IOException e) {
	  e.printStackTrace();
	 }
 * </pre>
 */
public class ExcelReader {

	private static ExcelReader instance = new ExcelReader();

	/**
	 *
	 * @return ExcelReader
	 */
	public static ExcelReader getInstance() {
		return instance;
	}

	/**
	 * .xls확장자파일
	 *
	 * @param File
	 * @return void
	 * @throws IOException
	 */
	public ArrayList<HashMap<Integer, Object>> readXls(File input) throws IOException {
		return readXls(new FileInputStream(input));
	}

	/**
	 * .xls확장자파일
	 *
	 * @param InputStream
	 * @return void
	 * @throws IOException
	 */
	public ArrayList<HashMap<Integer, Object>> readXls(InputStream input) throws IOException {
		FileInputStream fis = (FileInputStream) input;
		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		HSSFSheet sheet = workbook.getSheetAt(0);
		ArrayList<HashMap<Integer, Object>> result = new ArrayList<>();
		for (int rowindex = 1; rowindex < sheet.getPhysicalNumberOfRows(); rowindex++) {
			HSSFRow row = sheet.getRow(rowindex);
			HashMap<Integer, Object> map = new HashMap<>();
			if (row != null) {
				for (int columnindex = 0; columnindex <= row.getPhysicalNumberOfCells(); columnindex++) {
					HSSFCell cell = row.getCell(columnindex);
					if (cell != null)
						map.put(columnindex, cell.getCellFormula().toString());
				}
			}
			result.add(map);
		}
		return result;

	}

	/**
	 * .xlsx확장자파일
	 *
	 * @param File
	 * @return void
	 * @throws IOException
	 */
	public ArrayList<HashMap<Integer, Object>> readXlsx(File input) throws IOException {
		return readXlsx(new FileInputStream(input));
	}

	/**
	 * .xlsx확장자파일
	 *
	 * @param InputStream
	 * @return void
	 * @throws IOException
	 */
	public ArrayList<HashMap<Integer, Object>> readXlsx(InputStream input) throws IOException {
		FileInputStream fis = (FileInputStream) input;
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);

		printXSSFSheet(sheet);

		ArrayList<HashMap<Integer, Object>> result = new ArrayList<>();
		for (int rowindex = 1; rowindex < sheet.getPhysicalNumberOfRows(); rowindex++) {
			XSSFRow row = sheet.getRow(rowindex);
			HashMap<Integer, Object> map = new HashMap<>();
			if (row != null) {
				for (int columnindex = 0; columnindex <= row.getPhysicalNumberOfCells(); columnindex++) {
					// 셀값을 읽는다
					XSSFCell cell = row.getCell(columnindex);
					if (cell != null)
						map.put(columnindex, cell.getRawValue().toString());
				}
				result.add(map);
			}
		}
		return result;
	}

	public void printHSSFSheet(HSSFSheet sheet) {
		int rowindex = 0;
		int columnindex = 0;
		// 행의 수
		int rows = sheet.getPhysicalNumberOfRows();
		for (rowindex = 1; rowindex < rows; rowindex++) {
			// 행을 읽는다
			HSSFRow row = sheet.getRow(rowindex);
			if (row != null) {
				// 셀의 수
				int cells = row.getPhysicalNumberOfCells();
				for (columnindex = 0; columnindex <= cells; columnindex++) {
					// 셀값을 읽는다
					HSSFCell cell = row.getCell(columnindex);
					String value = "";
					// 셀이 빈값일경우를 위한 널체크
					if (cell == null) {
						continue;
					} else {
						// 타입별로 내용 읽기
						switch (cell.getCellType()) {
						case HSSFCell.CELL_TYPE_FORMULA:
							value = cell.getCellFormula();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:
							value = cell.getNumericCellValue() + "";
							break;
						case HSSFCell.CELL_TYPE_STRING:
							value = cell.getStringCellValue() + "";
							break;
						case HSSFCell.CELL_TYPE_BLANK:
							value = cell.getBooleanCellValue() + "";
							break;
						case HSSFCell.CELL_TYPE_ERROR:
							value = cell.getErrorCellValue() + "";
							break;
						}
					}
				}
			}
		}
	}

	/**
	 *
	 * @return void
	 * @throws IOException
	 */
	public void printXSSFSheet(XSSFSheet sheet) {
		int rowindex = 0;
		int columnindex = 0;
		// 시트 수 (첫번째에만 존재하므로 0을 준다)
		// 만약 각 시트를 읽기위해서는 FOR문을 한번더 돌려준다

		// 행의 수
		int rows = sheet.getPhysicalNumberOfRows();
		for (rowindex = 1; rowindex < rows; rowindex++) {
			// 행을읽는다
			XSSFRow row = sheet.getRow(rowindex);
			if (row != null) {
				// 셀의 수
				int cells = row.getPhysicalNumberOfCells();
				for (columnindex = 0; columnindex <= cells; columnindex++) {
					// 셀값을 읽는다
					XSSFCell cell = row.getCell(columnindex);
					String value = "";
					// 셀이 빈값일경우를 위한 널체크
					if (cell == null) {
						continue;
					} else {
						// 타입별로 내용 읽기
						switch (cell.getCellType()) {
						case XSSFCell.CELL_TYPE_FORMULA:
							value = cell.getCellFormula();
							break;
						case XSSFCell.CELL_TYPE_NUMERIC:
							value = cell.getNumericCellValue() + "";
							break;
						case XSSFCell.CELL_TYPE_STRING:
							value = cell.getStringCellValue() + "";
							break;
						case XSSFCell.CELL_TYPE_BLANK:
							value = cell.getBooleanCellValue() + "";
							break;
						case XSSFCell.CELL_TYPE_ERROR:
							value = cell.getErrorCellValue() + "";
							break;
						}
					}
				}
			}
		}
	}
}
