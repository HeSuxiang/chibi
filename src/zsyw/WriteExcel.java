package zsyw;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	private static final String EXCEL_XLS = "xls";
	private static final String EXCEL_XLSX = "xlsx";

	private static final String[] listHead = {
			"姓名",	
			"周日加班工时/H",	
			"请假工时/H",	
			"调休工时/H",	
			//"请病假工时/H",	
			"忘记打卡"
			
	};
	
	/**
	 * 从excel读取数据/往excel中写入 excel有表头。表头每列的内容相应实体类的属性
	 * 
	 * @author nagsh
	 * 
	 */
	public static   class ExcelManage {
		private static HSSFWorkbook workbook = null;
		
		/**
		 * 推断文件是否存在.
		 * @param fileDir  文件路径
		 * @return
		 */
		public static  boolean createFile(String fileDir){
			 boolean flag = false;
			 File file = new File(fileDir);
			 if(! file.exists()) {
				try {
					file.createNewFile();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} 
			 }
			 return flag;
		}
		/**
		 * 推断文件的sheet是否存在.
		 * @param fileDir   文件路径
		 * @param sheetName  表格索引名
		 * @return
		 */
		public static boolean createSheet(String fileDir,String sheetName){
			 boolean flag = false;
			 File file = new File(fileDir);
			 if(file.exists()){    //文件存在
	 			//创建workbook
	 	    	 try {
					workbook = new HSSFWorkbook(new FileInputStream(file));
					//加入Worksheet（不加入sheet时生成的xls文件打开时会报错)
		 	    	HSSFSheet sheet = workbook.getSheet(sheetName);  
		 	    	if(sheet!=null)
		 	    		flag = true;
				} catch (Exception e) {
					e.printStackTrace();
				} 
	 	    	
			 }else{    //文件不存在
				 flag = false;
			 }
			 
			 return flag;
		}
		/**
		 * 创建新excel.
		 * @param fileDir  excel的路径
		 * @param sheetName 要创建的表格索引
		 * @param titleRow excel的第一行即表格头
		 */
	    public static void createExcel(String fileDir,String sheetName) {
	    //,String titleRow[]){
	    	//创建workbook
	    	workbook = new HSSFWorkbook();
	    	//加入Worksheet（不加入sheet时生成的xls文件打开时会报错)
	    	Sheet sheet1 = workbook.createSheet(sheetName);  
	    	//新建文件
	    	FileOutputStream out = null;
	    	try {
				//加入表头
//		    	Row row = workbook.getSheet(sheetName).createRow(0);    //创建第一行  
//		    	for(int i = 0;i < titleRow.length;i++){
//		    		Cell cell = row.createCell(i);
//		    		cell.setCellValue(titleRow[i]);
//		    	}
		    	
		    	out = new FileOutputStream(fileDir);
				workbook.write(out);
			} catch (Exception e) {
				e.printStackTrace();
			} finally {  
			    try {  
			        out.close();  
			    } catch (IOException e) {  
			        e.printStackTrace();
			    }  
			}  
	    

	    }
	    /**
	     * 删除文件.
	     * @param fileDir  文件路径
	     */
	    public static boolean deleteExcel(String fileDir){
	    	boolean flag = false;
	    	File file = new File(fileDir);
	    	// 推断文件夹或文件是否存在  
	        if (!file.exists()) {  // 不存在返回 false  
	            return flag;  
	        } else {  
	            // 推断是否为文件  
	            if (file.isFile()) {  // 为文件时调用删除文件方法  
	                file.delete();
	                flag = true;
	            } 
	        }
	        return flag;
	    }
	    
	}
	
	
	
	public static void writeExcel(List<kaoqingInput> dataList, int cloumnCount, String finalXlsxPath) throws IOException {
		OutputStream out = null;
		//try {
			
			  ExcelManage.deleteExcel(finalXlsxPath);
			  ExcelManage.createFile(finalXlsxPath);
			  ExcelManage.createExcel(finalXlsxPath, "sheet");
			  ExcelManage.createSheet(finalXlsxPath, "sheet");
			  
			
			// 获取总列数
			int columnNumCount = cloumnCount;
			// 读取Excel文档
			File finalXlsxFile = new File(finalXlsxPath);
			
			//创建excel文件
			//File file=new File("f://poi.xlsx");
//			try {
//				finalXlsxFile.createNewFile();
//			    //将excel写入
//			    FileOutputStream stream= FileUtils.openOutputStream(finalXlsxFile);
//			    workbook.write(stream);
//			    stream.close();
//			} catch (IOException e) {
//			    e.printStackTrace();
//			    
//			if(!finalXlsxFile.exists()) {
//				finalXlsxFile.createNewFile();
//				
//			}
			
			Workbook workBook = getWorkbok(finalXlsxFile);
			// sheet 对应一个工作页
			
			workBook.removeSheetAt(0);
			workBook.createSheet();
			Sheet sheet = workBook.getSheetAt(0);
//			/**
//			 * 删除原有数据，除了属性列
//			 */
//			int rowNumber = sheet.getLastRowNum(); // 第一行从0开始算
//			System.out.println("原始数据总行数，除属性列：" + rowNumber);
//			for (int i = 1; i <= rowNumber; i++) {
//				Row row = sheet.getRow(i);
//				sheet.removeRow(row);
//			}
			// 创建文件输出流，输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
			out = new FileOutputStream(finalXlsxPath);
			workBook.write(out);
			/**
			 * 往Excel中写新数据
			 */
			Row rowHead = sheet.createRow(0);
			for (int i = 0; i < listHead.length; i++) {
				
				Cell cell = rowHead.createCell(i);
				cell.setCellValue(listHead[i]);
//				// 在一行内循环
//				Cell first = row.createCell(0);
//				first.setCellValue(name);
//
//				Cell second = row.createCell(1);
//				second.setCellValue(address);
//
//				Cell third = row.createCell(2);
//				third.setCellValue(phone);
 			}
				
				
			for (int j = 0; j < dataList.size(); j++) {
				// 创建一行：从第二行开始，跳过属性列
				Row row = sheet.createRow(j + 1);
				// 得到要插入的每一条记录
 				kaoqingInput item   = dataList.get(j);
//				String name = dataMap.get("BankName").toString();
//				String address = dataMap.get("Addr").toString();
//				String phone = dataMap.get("Phone").toString();
				for (int k = 0; k <= columnNumCount; k++) {
					// 在一行内循环
					Cell first = row.createCell(0);
					first.setCellValue(item.getName());

					Cell second = row.createCell(1);
					second.setCellValue(item.getZhourijiabangongshi());

					Cell third = row.createCell(2);
					third.setCellValue(item.getQingjiagongshi());
					
					Cell four = row.createCell(3);
					four.setCellValue(item.getTiaoxiugongshi());
					
					Cell five = row.createCell(4);
					five.setCellValue(item.getWanjidaka());
				}
			}
			// 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
			out = new FileOutputStream(finalXlsxPath);
			workBook.write(out);
//		} catch (Exception e) {
//			e.printStackTrace();
//		} finally {
//			try {
//				if (out != null) {
//					out.flush();
//					out.close();
//				}
//			} catch (IOException e) {
//				e.printStackTrace();
//			}
//		}
		System.out.println("数据导出成功");
		
		
	}

	/**
	 * 判断Excel的版本,获取Workbook
	 * 
	 * @param in
	 * @param filename
	 * @return
	 * @throws IOException
	 */
	public static Workbook getWorkbok(File file) throws IOException {
		Workbook wb = null;
		FileInputStream in = new FileInputStream(file);
		if (file.getName().endsWith(EXCEL_XLS)) { // Excel 2003
			wb = new HSSFWorkbook(in);
			 
		} else if (file.getName().endsWith(EXCEL_XLSX)) { // Excel 2007/2010
			wb = new XSSFWorkbook(in);
		}
		return wb;
	}
}
