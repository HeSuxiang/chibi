package zsyw;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import javax.swing.event.InternalFrameEvent;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadKaoqing {
	private static String FilePath = "D:\\用户目录\\我的文档\\chibi\\201808考勤信息明细表deleted.xls";
	private static final String EXCEL_XLS = "xls";
	private static final String EXCEL_XLSX = "xlsx";

	private static ArrayList<kaoqingInput> list = new ArrayList<kaoqingInput>();

	// �ж�Excel�İ汾,��ȡWorkbook
	public static Workbook getWorkbok(InputStream in, File file) throws IOException {
		Workbook wb = null;
		if (file.getName().endsWith(EXCEL_XLS)) { // Excel 2003
			wb = new HSSFWorkbook(in);
		} else if (file.getName().endsWith(EXCEL_XLSX)) { // Excel 2007/2010
			wb = new XSSFWorkbook(in);
		}
		return wb;
	}

	// �ж��ļ��Ƿ���excel
	public static void checkExcelVaild(File file) throws Exception {
		if (!file.exists()) {
			throw new Exception("文件不存在");
		}
		if (!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))) {
			throw new Exception("文件不是Excel");
		}
	}

	// ��ȡcell value
	public static String getCellValue(Cell cell) {

		SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
		int cellType = cell.getCellType();
		String cellValue = "";
		switch (cellType) {
		case Cell.CELL_TYPE_STRING: // 文本
			cellValue = cell.getRichStringCellValue().getString();
			break;
		case Cell.CELL_TYPE_NUMERIC: // 数字、日期
			if (DateUtil.isCellDateFormatted(cell)) {
				cellValue = fmt.format(cell.getDateCellValue());
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cellValue = String.valueOf(cell.getRichStringCellValue().getString());
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 布尔型
			cellValue = String.valueOf(cell.getBooleanCellValue()) + "#";
			break;
		case Cell.CELL_TYPE_BLANK: // 空白
			cellValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_ERROR: // 错误
			cellValue = "����#";
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			// 得到对应单元格的公式
			// cellValue = cell.getCellFormula() + "#";
			// 得到对应单元格的字符串
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cellValue = String.valueOf(cell.getRichStringCellValue().getString()) + "#";
			break;
		default:
			cellValue = "#";
		}

//		System.out.print(cellValue);

		if (cellValue.equals("")) {
			return "0";
		}
		return cellValue;
	}

	// Sheet List
	public static ArrayList<kaoqingInput> exportListFromExcel(String filePath) throws IOException {

		// try {
		// 同时支持Excel 2003、2007
		FilePath = filePath;
		File excelFile = new File(FilePath); //
		FileInputStream is = new FileInputStream(excelFile); //

		try {
			checkExcelVaild(excelFile);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		Workbook workbook = getWorkbok(is, excelFile);
		// Workbook workbook = WorkbookFactory.create(is); // // 这种方式
		// Excel2003/2007/2010都是可以处理的

		int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量
		/**
		 * 设置当前excel中sheet的下标：0开始
		 */
		Sheet sheet = workbook.getSheetAt(0); // 遍历第一个Sheet

//              // 为跳过第一行目录设置count  
//            for(int i = 1; i < sheetCount; i++) {
//            	String sheetName= sheet.getSheetName();
//            	if(sheetName.equals("����")) {
//            		break;
//            	}
//            		
//            	 
//            	sheet = workbook.getSheetAt(i);
//            }

		list = new ArrayList<kaoqingInput>();

		// 为跳过第一行目录设置count
		int count = 0;

		for (Row row : sheet) {
			// 跳过第N行
			if (count == 0 || count == 1) {
				count++;
				continue;
			}
			System.out.println("第" + count + "行");
			count++;

			// 如果当前行没有数据，跳出循环
			if (row.getCell(2) == null) {
				return list;
			}
			if (row.getCell(2).toString().equals("")) {
				return list;
			}

			if (row.getCell(2).toString().isEmpty()) {
				return list;
			}
			// String rowValue = ""; //

			// String cellValue = "";

			// 如果当前行没有31号, 算到30号
			int col = 31;
			if (row.getCell(95) == null) {
				col = 30;
			}
			else if (row.getCell(95).toString().equals("")) {
				col = 30;
			}

			else if (row.getCell(95).toString().isEmpty()) {
				col = 30;
			}
			
			else if (row.getCell(96) == null) {
				col = 30;
			}
			else if (row.getCell(96).toString().equals("")) {
				col = 30;
			}

			else if (row.getCell(96).toString().isEmpty()) {
				col = 30;
			}
			

			String name = getCellValue(row.getCell(2));

			System.out.println("姓名" + name);
			kaoqingInput item = new kaoqingInput(name);
			int base = 2;

			// everyday
			for (int i = 0; i < col; i++) {
				System.out.println(name + "开始 i=" + i);

				String cellValue = getCellValue(row.getCell((i * 3) + base + 3));
//				if(cellValue == null)
//				{
//					System.out.println("cellValue == null" );
//					break;
//				}else {
//					System.out.println("cellValue != null" );
//				}

				if (cellValue.indexOf("忘记打卡") != -1) {
					item.AddWanjidaka();
				}

				int qingjiaIndex = cellValue.indexOf("请假");
				if (qingjiaIndex != -1) {
					int index = cellValue.indexOf("H", qingjiaIndex);
					String qingjia = cellValue.substring(qingjiaIndex + 2, index);
					System.out.println(name + "  qingjia i =" + i + " " + qingjia);
					item.AddQingjiagongshi(Float.parseFloat(qingjia));
				}

				int jibanIndex = cellValue.indexOf("加班");
				if (jibanIndex != -1) {
					int index = cellValue.indexOf("H", jibanIndex);
					String jiban = cellValue.substring(jibanIndex + 2, index);
					System.out.println(name + "  jiban i =" + i + " " + jiban);
					item.AddZhourijiabangongshi(Float.parseFloat(jiban));

				}

				int tiaoxiuIndex = cellValue.indexOf("调休");
				if (tiaoxiuIndex != -1) {
					int index = cellValue.indexOf("H", tiaoxiuIndex);
					String tiaoxiu = cellValue.substring(tiaoxiuIndex + 2, index);
					System.out.println(name + "  tiaoxiu i =" + i + " " + tiaoxiu);
					item.AddTiaoxiugongshi(Float.parseFloat(tiaoxiu));
				}

//				else if(cellValue.startsWith("请病假")) {
//					int index = cellValue.indexOf("H");
//					cellValue.indexOf(str)
//					String qingjia = cellValue.substring(3, index);
//					System.out.println(name + "  请病假 i =" +i + " " +qingjia);
//					 
//				}

				// System.out.println("wanjidaka i =" +i + " " +wanjidaka);
			}
			// System.out.println(name + " wanjidaka =" + item.getWanjidaka());
			// String type = getCellValue(row.getCell(4));
			// int ChuhuoCount = Integer.parseInt( getCellValue(row.getCell(5)) );

			list.add(item);

			// System.out.println(item.getName() +" 忘记打卡共计" + item.getWanjidaka()+"次");

//                for (Cell cell : row) {  
//                	if(cell.toString() == null){  
//                		continue;  
//                	}  
//
//
//                	rowValue += cellValue;  
//                }    
			// System.out.println(rowValue);
			// System.out.println("" + count);
		}
		// } catch (Exception e) {
		// e.printStackTrace();
		// } finally{

		return list;
		// }
	}
}