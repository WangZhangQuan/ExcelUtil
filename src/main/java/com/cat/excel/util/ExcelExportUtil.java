package com.cat.excel.util;

import java.io.*;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.AbstractCollection;
import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.hutool.core.bean.BeanResolver;

/**
 * @author Cat.Wang
 *
 */
public class ExcelExportUtil {

	private static HashMap<CellStyle, CellStyle> CELL_STYLE_CACHE = new HashMap<CellStyle, CellStyle>();
	
	private Object model = null;
	private String excelFilePath = null;
	private InputStream inputStream = null;
	private Workbook book = null;
	private HSSFWorkbook hBook = null;
	private XSSFWorkbook xBook = null;
	
	/**
	 * 当前扩展行的开始位置
	 */
	private Integer csr = 0;
	/**
	 * 当前扩展行的结束位置
	 */
	private Integer cer = 0;
	/**
	 * 当前扩展单元格的合并区域
	 */
	private CellRangeAddress cr = null;
	/**
	 * 当前扩展单元格的行的最大行合并区域
	 */
	private CellRangeAddress crx = null;
	/**
	 * 当前单元格的合并行个数
	 */
	private Integer co = 1;
	/**
	 * 当前单元格的行的最大合并行个数
	 */
	private Integer cox = 1;
	/**
	 * 本次打印行的个数
	 */
	private Integer cn = 0;
	/**
	 * 本次扩展行的个数
	 */
	private Integer cexr = 0;
	/**
	 * 当前操作的单元格
	 */
	private Cell cc = null;
	/**
	 * 当前操作的模板单元格
	 */
	private Cell ctpl = null;
	/**
	 * 当前操作的sheet名称
	 */
	private String csn = null;
	
	/**
	 * 记录合并单元格已被处理过的集合
	 */
	private Set<CellRangeAddress> proccessedCas = new HashSet<CellRangeAddress>();
	
	/**
	 * 存放着所有sheet的rowRanges
	 */
	private Map<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>> sheetsRowRanges = new HashMap<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>>();
	
	/**
	 * 存放着所有sheet的cellRanges
	 */
	private Map<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>> sheetsCellRanges = new HashMap<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>>();
	
	/**
	 * Map<递归深度, ArrayList<KeyValue<扩展时的表达式存在的单元格, HashMap<开始的行 闭区间, 结束的行 闭区间>>>>
	 * 记录打印多行数据的表达式扩展的行
	 */
	private HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>> rowRanges = null;
	/**
	 * Map<递归深度, ArrayList<KeyValue<扩展时的表达式存在的单元格, HashMap<开始的列 闭区间, 结束的列 闭区间>>>>
	 * 记录打印多列数据的表达式扩展的列
	 */
	private HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>> cellRanges = null;

	protected ExcelExportUtil() {
		CELL_STYLE_CACHE.clear();
	}

	private void setRangesBySheetName(String sheetName) {
		if(!sheetsRowRanges.containsKey(sheetName)) {
			rowRanges = new HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>();
			sheetsRowRanges.put(sheetName, rowRanges);
		}else {
			rowRanges = sheetsRowRanges.get(sheetName);
		}
		if(!sheetsCellRanges.containsKey(sheetName)) {
			cellRanges = new HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>();
			sheetsCellRanges.put(sheetName, cellRanges);
		}else {
			cellRanges = sheetsCellRanges.get(sheetName);
		}
	}
	
	/**
	 * 获得工具实例
	 * 
	 * @param model 
	 * @param excelFilePath
	 * @return
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static ExcelExportUtil createInstance(Object model, String excelFilePath)
			throws InvalidFormatException, IOException {
		ExcelExportUtil excelExportUtil = new ExcelExportUtil();
		excelExportUtil.model = model;
		excelExportUtil.excelFilePath = excelFilePath;
		excelExportUtil.inputStream = new FileInputStream(excelExportUtil.excelFilePath);

		if (!excelExportUtil.inputStream.markSupported()) {
			excelExportUtil.inputStream = new PushbackInputStream(excelExportUtil.inputStream, 8);
		}

		if (POIFSFileSystem.hasPOIFSHeader(excelExportUtil.inputStream)) {
			excelExportUtil.hBook = new HSSFWorkbook(excelExportUtil.inputStream);
			excelExportUtil.book = excelExportUtil.hBook;
		} else if (POIXMLDocument.hasOOXMLHeader(excelExportUtil.inputStream)) {
			excelExportUtil.xBook = new XSSFWorkbook(OPCPackage.open(excelExportUtil.inputStream));
			excelExportUtil.book = excelExportUtil.xBook;
		}

		return excelExportUtil;
	}
	
	public static ExcelExportUtil createInstance(Object model, InputStream excelFileStream)
			throws InvalidFormatException, IOException {
		ExcelExportUtil excelExportUtil = new ExcelExportUtil();
		excelExportUtil.model = model;
		excelExportUtil.inputStream = excelFileStream;

		if (!excelExportUtil.inputStream.markSupported()) {
			excelExportUtil.inputStream = new PushbackInputStream(excelExportUtil.inputStream, 8);
		}

		if (POIFSFileSystem.hasPOIFSHeader(excelExportUtil.inputStream)) {
			excelExportUtil.hBook = new HSSFWorkbook(excelExportUtil.inputStream);
			excelExportUtil.book = excelExportUtil.hBook;
		} else if (POIXMLDocument.hasOOXMLHeader(excelExportUtil.inputStream)) {
			excelExportUtil.xBook = new XSSFWorkbook(OPCPackage.open(excelExportUtil.inputStream));
			excelExportUtil.book = excelExportUtil.xBook;
		}

		return excelExportUtil;
	}
	
	public void writeBook(String file) throws IOException {

		if (gethBook() != null) {
			gethBook().write(new FileOutputStream(file));
		} else {
			getxBook().write(new FileOutputStream(file));
		}

	}

	public static final String EXP_START = "_Exp_{";
	public static final String EXP_END = "}";
	
	
	/**
	 * 通过sheet 获取sheet的model
	 * @param sheet
	 * @return
	 */
	public Object getSheetModel(Sheet sheet) {
		return getValueByExpressionAndModel(EXP_START + sheet.getSheetName() + EXP_END, this.model);
	}
	
	/**
	 * 装配模板
	 * @return
	 */
	public String parse_0_1() {
		StringBuilder message = new StringBuilder();
		
		// 遍历所有sheet
		for (int i = 0; i < this.book.getNumberOfSheets(); i++) {
			
			String ssn = this.book.getSheetName(i);
			
			List<String> sheetNames = setSheetNamesByExpression(i, ssn);
			
			for (int x = 0; x < sheetNames.size(); x++) {
				
				this.csn = sheetNames.get(x);
				
				Sheet sheet = this.book.getSheet(this.csn);

				// 切换ranges环境
				setRangesBySheetName(this.csn);
				
				// 遍历所有行
				for (int j = 0; j <= sheet.getLastRowNum(); j++) {
					
					Row row = sheet.getRow(j);
					
					if(row != null) {
						// 遍历所有单元格
						for (int k = 0; k <= row.getLastCellNum(); k++) {
							Cell cell = row.getCell(k);
							
							if(cell != null) {
								// 判断是否是合并单元格 
								if(isMergedRegion(sheet, j, k) && !isCellContainsProccessedCas(j, k)) { // 判断此合并单元格未被处理过
									String value = getMergedRegionValue(sheet, j, k); // 获取合并单元格未被处理过的值
									setCellValueByExpressionAndModel(value, getSheetModel(sheet), this.book, sheet, row, cell); // 通过表达式设置值
								}else {
									String value = getCellValue(cell); // 获取单元格的值
									setCellValueByExpressionAndModel(value, getSheetModel(sheet), this.book, sheet, row, cell); // 通过表达式设置值
								}
							}
							
						}
					}
					
				}
			}
			
		}
		return message.toString();
	}
	
	private List<String> setSheetNamesByExpression(Integer sheetIndex, Object expr) {
		Object v = expr;
		// 如果是表达式则取表达式值
		if(isExpression(v)) {
			v = getValueByExpressionAndModel(v.toString(), this.model);
			return setSheetNamesByExpression(sheetIndex, v);
		}else {
			List<String> ssns = new ArrayList<String>(1);
			ssns.addAll(setSheetName(sheetIndex, v));
			return ssns;
		}
	}
	
	private List<String> setSheetName(Integer sheetIndex, Object v) {
		
		List<String> ssns = new ArrayList<String>(1);
		
		if(v != null) {
			// 如果sheet名称是集合类型
			if(v instanceof AbstractCollection<?>) {
				int i = 0;
				AbstractCollection<?> l = ((AbstractCollection<?>)v);
				ssns = new ArrayList<String>(l.size());
				for (Iterator<?> iterator = l.iterator(); iterator.hasNext(); i++) {
					
					Object o = (Object) iterator.next();
					Integer csi = null;
					// 将第一个表达式sheet利用起来
					if(i == 0) {
						csi = sheetIndex;
					}else {
						Sheet sheet = this.book.cloneSheet(sheetIndex);
						csi = this.book.getSheetIndex(sheet);
					}
					String sheetName = null;
					if(o != null) {
						sheetName = o.toString();
					}
					
					ssns.add(setReallySheetName(csi, sheetName));
				}
				if(l.size() == 0) {
					ssns.add(setReallySheetName(sheetIndex, "没有任何数据"));
				}
			}else if(v instanceof AbstractMap<?, ?>) {
				throw new RuntimeException("暂不支持Map类型设置为sheetName");
			}else {
				ssns.add(setReallySheetName(sheetIndex, v.toString()));
			}
		}
		
		return ssns;
	}
	
	private String setReallySheetName(Integer sheetIndex, String sheetNamex) {
		
		String sheetName = "Sheet";
		
		if(sheetNamex != null && validateSheetName(sheetNamex)) {
			sheetName = sheetNamex;
		}
		
		String name = sheetName;
		Sheet sheet = null;
		// 如果同名工作表不是当前sheet则重新生成名称
		for(int i = 1; (sheet = this.book.getSheet(name)) != null && this.book.getSheetIndex(sheet) != sheetIndex; i++) {
			name = sheetName + i;
		}
		
		if(validateSheetName(name)) {
			this.book.setSheetName(sheetIndex, name);
			sheetName = name;
		}
		
		return sheetName;
	}
	
	private boolean validateSheetName(String name) {
		if(name == null || "".equals(name)) {
			return false;
		}
		if(name.startsWith("'")) {
			return false;
		}
		if(name.length() > 31) {
			return false;
		}
		if(name.matches("[\\\\/\\?\\*\\[\\]]+")) {
			return false;
		}
		
		return true;
	}
	
	/**
	 * 将model数据填充到Excel
	 * 此方法bug太多 请使用parseNew方法
	 * @return
	 */
	@Deprecated
	public String parse() {
		StringBuilder message = new StringBuilder();
		
		
		
		
		@SuppressWarnings("unused")
		Map<Integer,Integer> ranges = new HashMap<Integer, Integer>();
		
		Sheet sheet = this.book.getSheetAt(0);

		for (int i = 0; i < sheet.getLastRowNum() + 1; ++i) {
			
			Row row = sheet.getRow(i);
			if (row != null) {
			
//				List<Integer> igColumns = new ArrayList<Integer>();
				int endRow = row.getRowNum();
				
				for (Cell c : row) {
					// 判断是否具有合并单元格
					if (isMergedRegion(sheet, i, c.getColumnIndex())) {
						if(!this.isCellContainsProccessedCas(row.getRowNum(), c.getColumnIndex())){
							String value = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
							// 得到一个表达式
							if(StringUtils.isNotBlank(value) && value.startsWith(EXP_START) && value.endsWith(EXP_END)){
								List<Object> list = this.executeExpress(clearExpressF(value), model);
								if(list != null){
									if(list.size() == 1){
										setCellValue(c, list.get(0));
									}else if(list.size() > 1){
										// 在下方插入行 其它单元格跨行
										setCellValue(c, list.get(0)); // 设置第一行的值
										if ((list.size() - (endRow - i) - 1) > 0) {
//											ranges.put(row.getRowNum(), list.size() - (endRow - i) - 1); // 放入扩充行
											sheet.shiftRows(endRow + 1, sheet.getLastRowNum(), list.size() - (endRow - i) - 1);
										}
										for(int x = 1; x < list.size(); ++x){
											Row r = sheet.getRow(i + x);
											if(endRow < i + x){
												endRow = x + i;
												r = sheet.createRow(i + x);
											}
											for(int z = 0; z < row.getLastCellNum();++z){
												Cell ce = r.getCell(z);
												if(ce == null){
													ce = r.createCell(z);
//													if(row.getCell(z) != null && sheet.getRow(i + x - 1).getCell(z) != null){
//														ce.setCellStyle(row.getCell(z).getCellStyle());
//														ce.setCellComment(row.getCell(z).getCellComment());
//														ce.setCellType(row.getCell(z).getCellType());
//														CellRangeAddress cra = this.getMergedRegion(sheet, row.getRowNum(), row.getCell(z).getColumnIndex()).copy();
//														Integer re = cra.getLastRow() - cra.getFirstRow();
//														cra.setFirstRow(ce.getRowIndex());
//														cra.setLastRow(ce.getRowIndex() + re);
//														sheet.addMergedRegion(cra);
//													}
												}
												if(ce.getColumnIndex() == c.getColumnIndex()){
													setCellValue(ce, list.get(x));
												}
											}
										}
//										igColumns.add(c.getColumnIndex()); // 忽略列
										
									}
								}
							}
						}
					} else {
						String value = getCellValue(c);
						// 得到一个表达式
						if(StringUtils.isNotBlank(value) && value.startsWith(EXP_START) && value.endsWith(EXP_END)){
							List<Object> list = this.executeExpress(clearExpressF(value), model);
							if(list != null){
								if(list.size() == 1){
									setCellValue(c, list.get(0));
								}else if(list.size() > 1){
									// 在下方插入行 其它单元格跨行
									setCellValue(c, list.get(0)); // 设置第一行的值
									if (sheet.getRow(endRow + 1) != null && (list.size() - (endRow - i) - 1) > 0) {
//										ranges.put(row.getRowNum(), list.size() - (endRow - i) - 1); // 放入扩充行
										sheet.shiftRows(endRow + 1, sheet.getLastRowNum(), list.size() - (endRow - i) - 1);
									}
									for(int x = 1; x < list.size(); ++x){
										Row r = sheet.getRow(i + x);
										if(endRow < i + x){
											endRow = x + i;
											r = sheet.createRow(i + x);
										}
										for(int z = 0; z < row.getLastCellNum();++z){
											Cell ce = r.getCell(z);
											if(ce == null){
												ce = r.createCell(z);
											}
											if(ce.getColumnIndex() == c.getColumnIndex()){
												setCellValue(ce, list.get(x));
											}
										}
									}
//									igColumns.add(c.getColumnIndex()); // 忽略列
									
								}
							}
						}
					}
				}
				

//				for(Cell c1 : row){ // 跨行未忽略的行
//					if(igCellList.contains(c1.getColumnIndex())){
//						this.mergeRegion(sheet, i, endRow, c1.getColumnIndex(), c1.getColumnIndex());
//					}
//				}
				
//				for(Cell c : row){
//					if(!igColumns.contains(c.getColumnIndex())){ // 不在忽略里面添加跨行
//						this.mergeRegion(sheet, row.getRowNum(), endRow, c.getColumnIndex(), c.getColumnIndex());
//					}
//				}
			}

			
		}

		// for(int i=startReadLine; i<sheet.getLastRowNum()-tailLine+1; i++) {
		// row = sheet.getRow(i);
		// for(Cell c : row) {
		// boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
		// //判断是否具有合并单元格
		// if(isMerge) {
		// String rs = getMergedRegionValue(sheet, row.getRowNum(),
		// c.getColumnIndex());
		// System.out.print(rs + “ “);
		// }else {
		// System.out.print(c.getRichStringCellValue()+” “);
		// }
		// }
		// System.out.println();

		return message.toString();
	}

	private Object getValueByExpressionAndModel(String expr, Object model) {
		
		Object value = null;
		
		try {
			String es = clearExpressF(expr);
			value = BeanResolver.resolveBean(model, es); // 采用第三方工具类获取数据 能穿透 pojo map list
		}catch(Exception e) {
		}
		
		return value;
	}
	
	private void setCellValueByExpressionAndModel(
			String expr,
			Object model,
			Workbook book,
			Sheet sheet,
			Row row,
			Cell cell
			) {

		// 判断是否是表达式
		if(isExpression(expr)) {
			// 设置当前模板单元格
			this.ctpl = cell;
			setCellValueByUnknowValue(getValueByExpressionAndModel(expr, model), book, sheet, row, cell, 0);
		}
	}
	
	/**
	 * @param value
	 * @param book
	 * @param sheet
	 * @param row
	 * @param cell
	 * @param depth 递归调用的深度 可以传空
	 */
	@SuppressWarnings("unused")
	private void setCellValueByUnknowValue(
			Object value,
			Workbook book,
			Sheet sheet,
			Row row,
			Cell cell,
			Integer depth
			) {
		
		
		if(depth == null) {
			depth = 0;
		}
		Integer nextDepth = depth + 1;
		
		if(value instanceof AbstractCollection<?>) { // 判断是否是集合类型
			
			AbstractCollection<?> c = ((AbstractCollection<?>)value);
			
			if(c == null || c.size() == 0) {
				setCellValue(cell, null);
				return ;
			}
			if(c.size() == 1) { // 无需扩展行 直接设置值
				setCellValueByUnknowValue(
						c.iterator().next(), 
						cell.getSheet().getWorkbook(),
						cell.getSheet(), 
						cell.getRow(), 
						cell, 
						nextDepth
						);
				return ;
			}
			 // 扩展行
			shiftCollection(sheet, cell, c, depth);
			
			// 保存单元格引用防止递归过后位置改变
			List<KeyValue<Cell, Object>> cvs = new ArrayList<KeyValue<Cell, Object>>();
			
			// 填充值
			int i = 0;
			for (Iterator<?> iterator = c.iterator(); iterator.hasNext();) {
				Object object = (Object) iterator.next();
				
				cell = createCellAndCopyStyle(sheet, cell, row.getRowNum() + i, cell.getColumnIndex());
				
				cvs.add(new KeyValue<Cell, Object>(cell, object));

				i += this.cox;
			}
			
			for (KeyValue<Cell, Object> kv : cvs) {
				setCellValueByUnknowValue(
						kv.getValue(), 
						kv.getKey().getSheet().getWorkbook(),
						kv.getKey().getSheet(), 
						kv.getKey().getRow(), 
						kv.getKey(), 
						nextDepth
						);
			}
		}else if(value instanceof AbstractMap<?, ?>) { // 判断是否是map类型

			AbstractMap<?, ?> m = ((AbstractMap<?, ?>)value);
			
			if(m == null) {
				setCellValue(cell, null);
				return;
			}
			// 扩展列
			shiftMap(sheet, cell, m, depth);
			
			// 填充值
//			int i = 0;
//			int rn = row.getRowNum() + 1;
//			Set<?> keySet = m.keySet();
//			for (Iterator<?> iterator = keySet.iterator(); iterator.hasNext();) {
//				Object object = (Object) iterator.next();
//				// 设置key值
//				setCellValueByUnknowValue(object, book, sheet, row, row.getCell(cell.getColumnIndex() + i), nextDepth);
//				// 设置value值
//				setCellValueByUnknowValue(m.get(object), book, sheet, sheet.getRow(rn), sheet.getRow(rn).getCell(cell.getColumnIndex() + i), nextDepth);
//				i++;
//			}
			
		}else { // 默认为基础类型 
			setCellValue(cell, value);
		}
		
	}
	
	private List<KeyValue<Cell, HashMap<Integer, Integer>>> getRangesByCellAndDepth(Cell cell, Integer depth){
		
		ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>> al = null;
		
		if(this.rowRanges.containsKey(depth)) {
			al = this.rowRanges.get(depth);
		}else {
			al = new ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>();
			this.rowRanges.put(depth, al);
		}
		
		List<KeyValue<Cell, HashMap<Integer, Integer>>> ml = new ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>();
		
		for (int i = 0; i < al.size(); i++) {
			
			if(isSameRowByRegin(al.get(i).getKey(), cell)) {
				ml.add(al.get(i));
			}
			
		}
		
		return ml;
	}
	
	private boolean isSameRowByRegin(Cell c1, Cell c2) {
		
		CellRangeAddress msrr1 = getMaxSameRowCellRowRegin(c1.getSheet(), c1.getRowIndex());
		CellRangeAddress msrr2 = getMaxSameRowCellRowRegin(c2.getSheet(), c2.getRowIndex());
		
		if(msrr1 == null || msrr2 == null) {
			if(msrr1 != null) {
				return msrr1.getFirstRow() <= c2.getRowIndex() && msrr1.getLastRow() >= c2.getRowIndex();	
			}else if(msrr2 != null){
				return msrr2.getFirstRow() <= c1.getRowIndex() && msrr2.getLastRow() >= c1.getRowIndex();	
			}else {
				return c1.getRowIndex() == c2.getRowIndex();	
			}
		}
		
		return 
				msrr1.getFirstRow() == msrr2.getFirstRow() 
				&& msrr1.getLastRow() == msrr2.getLastRow()
				;
	}
	
	private void shiftCollection(Sheet sheet, Cell cell, AbstractCollection<?> c, Integer depth) {
		
		initShiftCollection(sheet, cell, c, depth);
		
		List<KeyValue<Cell, HashMap<Integer, Integer>>> ml = getRangesByCellAndDepth(cell, depth);
		
		if(ml.size() <= 0) { // 需要全部扩展
			// 扩展行
			addMergedRegionAndSheetRow(sheet);
			// 截断浅深度同行的range
			truncationRowsRange(sheet, depth);
			// 添加扩展记录
			addExtRange(sheet, depth);
		}else if(ml.size() > 0) { // 需要部分扩展
			Integer size = computeExtSize(ml); // 计算已经扩展的行数
			if(size < this.cer - (cell.getRowIndex() + this.cox)) { // 如果当前集合单元格个数大于已经扩展的单元格个数
				Integer maxRangeEnd = getMaxRangeEnd(ml);
				// 扩展行
				shiftRowsAndAvoidDepthRanges(sheet, maxRangeEnd, depth);
				// 截断浅深度同行的range
				truncationRowsRange(sheet, depth);
				// 添加扩展记录
				addExtRange(sheet, depth);
			}else {
				// 刚好满足无需扩展 只添加合并
				addMergedRegion(sheet);
			}
			
		}
		
	}
	
	private void initShiftCollection(Sheet sheet, Cell cell, AbstractCollection<?> c, Integer depth) {
		
		this.cc = cell;
		
		this.cn = c.size();
		this.cexr = this.cn - 1;
		
		this.cr = getMergedRegionByCell(sheet, cell.getRowIndex(), cell.getColumnIndex());
		
		if(depth <= 0) {
			this.crx = getMaxSameRowCellRowRegin(sheet, cell.getRowIndex());
		}else if(this.crx != null){
			CellRangeAddress msrcr = getMaxSameRowCellRowRegin(sheet, this.ctpl.getRowIndex());
			Integer ind = cell.getRowIndex() - (this.ctpl.getRowIndex() - msrcr.getFirstRow());
			this.crx = new CellRangeAddress(ind, ind + (msrcr.getLastRow() - msrcr.getFirstRow()), msrcr.getFirstColumn(), msrcr.getLastColumn());
		}
		
		this.co = 1;
		this.cox = 1;
		
		if(this.cr != null) {
			this.co = this.cr.getLastRow() - this.cr.getFirstRow() + 1;
		}
		
		Integer rowIndex = cell.getRowIndex();
		
		if(this.crx != null) {
			this.cox = this.crx.getLastRow() - this.crx.getFirstRow() + 1;
			rowIndex = this.crx.getFirstRow();
		}
		
		this.csr = rowIndex + this.cox;
		this.cer = rowIndex + this.cox + this.cexr * this.cox;
	}
	
	private void shiftRowsAndAvoidDepthRanges(Sheet sheet, Integer startRow, Integer depth) {
		
		Set<Integer> ks = this.rowRanges.keySet();
		for (Iterator<?> i = ks.iterator(); i.hasNext();) {
			Integer k = (Integer) i.next();
			if(k > depth) {
				List<KeyValue<Cell, HashMap<Integer, Integer>>> al = this.rowRanges.get(k);
				
				for (Iterator<KeyValue<Cell, HashMap<Integer, Integer>>> ix = al.iterator(); ix.hasNext();) {
					KeyValue<Cell, HashMap<Integer, Integer>> v = ix.next();
					// 如果最后一行是列表 则从列表的最后位置开始扩展
					if(v.getKey().getRowIndex() == (startRow - 1)) {
						for (Iterator<Integer> ixx = v.getValue().values().iterator(); ixx.hasNext();) {
							Integer vx = ixx.next();
							if(vx > startRow) {
								startRow = vx;
							}
						}
					}
				}
			}
		}
		
		this.csr = startRow;
		
		addMergedRegionAndSheetRow(sheet);
	}
	
	private Cell createCellAndCopyStyle(Sheet sheet, Cell template, Integer rowIndex, Integer columnIndex){
		
		Row row = sheet.getRow(rowIndex);
		
		if(row == null) {
			row = sheet.createRow(rowIndex);
		}
		
		Cell cell = row.getCell(columnIndex);
		
		if(cell == null) {
			cell = row.createCell(columnIndex);
		}
		
		if(template != null) {
			copyCell(sheet.getWorkbook(), template, cell, false);
		}
		
		return cell;
	}
	
	private void addMergedRegionAndSheetRow(Sheet sheet) {

		// 扩展行
		if(sheet.getLastRowNum() < this.csr) {
			for (int i = this.csr; i <= this.cer; i += this.cox) {
				sheet.createRow(i);
			}
		}else {
			sheet.shiftRows(this.csr, sheet.getLastRowNum(), this.cer - this.csr);
		}
		addMergedRegion(sheet);
	}
	
	private  void addMergedRegion(Sheet sheet) {
		
		// 添加合并单元格
		if(this.cr != null) {
			for (int i = this.cc.getRowIndex(); i < this.cer; i += (this.cox - 1)) {
				CellRangeAddress cra = new CellRangeAddress(i, i + (this.co - 1), this.cr.getFirstColumn(), this.cr.getLastColumn());
				sheet.addMergedRegion(cra);
				i += 1;
			}
		}
	}
	
	private CellRangeAddress getMaxSameRowCellRowRegin(Sheet sheet, Integer rowIndex) {
		CellRangeAddress r = null;
		Row row = sheet.getRow(rowIndex);
		if(row != null) {
			for (Cell cell : row) {
				if(cell != null) {
					CellRangeAddress range = getMergedRegionByCell(sheet, cell.getRowIndex(), cell.getColumnIndex());
					if(range != null) {
						if(r == null) {
							r = range;
						}else {
							if((r.getLastRow() - r.getFirstRow()) < (range.getLastRow() - range.getFirstRow())) {
								r = range;
							}
						}
					}
				}
			}
		}
		// 得到合成最大的合并范围
		r = getMultiMaxSameRowCellRowRegin(sheet, r);
		
		return r;
	}
	
	private CellRangeAddress getMultiMaxSameRowCellRowRegin(Sheet sheet, CellRangeAddress r) {
		if(r == null) {
			return null;
		}
		List<CellRangeAddress> arr = getAllRowRegin(sheet, r.getFirstRow(), r.getLastRow());
		
		Integer fri = r.getFirstRow();
		Integer lri = r.getLastRow();
		Integer fci = r.getFirstColumn();
		Integer lci = r.getLastColumn();
		
		for (int i = 0; i < arr.size(); i++) {
			CellRangeAddress cra = arr.get(i);
			if(cra.getFirstRow() < fri) {
				fri = cra.getFirstRow();
			}else if(cra.getLastRow() > lri) {
				lri = cra.getLastRow();
			}else if(cra.getFirstColumn() < fci) {
				fci = cra.getFirstColumn();
			}else if(cra.getLastColumn() > lci) {
				lci = cra.getLastColumn();
			}
		}
		
		return new CellRangeAddress(fri, lri, fci, lci);
	}
	
	private List<CellRangeAddress> getAllRowRegin(Sheet sheet, Integer firstRowIndex, Integer endRowIndex){

		List<CellRangeAddress> cras = new ArrayList<CellRangeAddress>();
		
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			
			// 相交或包含
			if(firstRow < firstRowIndex && lastRow > endRowIndex) {
				cras.add(range);
			}else if(firstRow > firstRowIndex && lastRow < endRowIndex) {
				cras.add(range);
			}else if(firstRow >= firstRowIndex && firstRow <= endRowIndex) {
				cras.add(range);
			}else if(lastRow <= endRowIndex && lastRow >= firstRowIndex) {
				cras.add(range);
			}
		}
		
		return cras;
	}
	
	@SuppressWarnings("unused")
	private Map<Integer, Integer> mergeMapRanges(List<HashMap<Integer, Integer>> maps){
		
		Map<Integer, Integer> mmr = new HashMap<Integer, Integer>();
		
		for (int i = 0; i < maps.size(); i++) {
			HashMap<Integer, Integer> m = maps.get(i);
			for (Iterator<?> ix = m.keySet().iterator(); ix.hasNext();) {
				Integer v = (Integer) ix.next();
				Integer vx = m.get(v);
				
				boolean flag = false;
				
				for (Iterator<?> ixx = mmr.keySet().iterator(); ixx.hasNext();) {
					Integer vxx = (Integer) ixx.next();
					Integer vxxx = mmr.get(vxx);
					
					if(v >= vxx && vx >= vxxx) {
						// 包含关系
						mmr.remove(vxx);
						mmr.put(v, vx);
						flag = true;
						break;
					}else if(v <= vxxx) {
						// 相交关系
						mmr.replace(vxx, vx);
						flag = true;
						break;
					}else if(vx >= vxx) {
						// 相交关系
						mmr.remove(vxx);
						mmr.put(v, vxxx);
						flag = true;
						break;
					}
				}
				
				// 并没合并map
				if(!flag) {
					// 相离关系
					mmr.put(v, vx);
				}
			}
		}
		
		return mmr;
	}
	
	@SuppressWarnings("unchecked")
	private Integer getMaxRangeEnd(List<KeyValue<Cell, HashMap<Integer, Integer>>> ml) {
		
		Integer max = 0;
		
		for (Iterator<?> i = ml.iterator(); i.hasNext();) {
			KeyValue<Cell, HashMap<Integer, Integer>> v = (KeyValue<Cell, HashMap<Integer, Integer>>) i.next();
			for (Iterator<?> ix = v.getValue().values().iterator(); ix.hasNext();) {
				Integer vx = (Integer) ix.next();
				if(vx > max) {
					max = vx;
				}
			}
		}
		
		return max;
	}
	
	private void addExtRange(Sheet sheet, Integer depth) {
		
		ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>> al = null;
		
		if(this.rowRanges.containsKey(depth)) {
			al = this.rowRanges.get(depth);
		}else {
			al = new ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>();
			this.rowRanges.put(depth, al);
		}
		
		HashMap<Integer, Integer> m = null;
		
		for (Iterator<?> i = al.iterator(); i.hasNext();) {
			@SuppressWarnings("unchecked")
			KeyValue<Cell, HashMap<Integer, Integer>> v = (KeyValue<Cell, HashMap<Integer, Integer>>) i.next();
			if(v.getKey().equals(this.cc)) {
				m = v.getValue();
				break;
			}
		}
		
		if(m == null) {
			m = new HashMap<Integer, Integer>();
		}
		
		m.put(this.cc.getRowIndex() + this.cox, this.cc.getRowIndex() + this.cox + this.cexr * this.cox);
		
		al.add(new KeyValue<Cell, HashMap<Integer, Integer>>(this.cc, m));
		
	}
	
	private void truncationRowsRange(Sheet sheet, Integer depth) {

		Iterator<Integer> i = this.rowRanges.keySet().iterator();
		
		while(i.hasNext()) {
			Integer k = i.next();
			if(k < depth) {
				List<KeyValue<Cell, HashMap<Integer, Integer>>> v = this.rowRanges.get(k);
				for (int j = 0; j < v.size(); j++) {
					HashMap<Integer, Integer> vx = v.get(j).getValue();
					for (Iterator<Integer> ix = vx.keySet().iterator(); ix.hasNext();) {
						Integer kx = ix.next();
						Integer vxx = vx.get(kx);
						
						if(kx <= this.cc.getRowIndex() + this.cox && vxx > this.cc.getRowIndex() + this.cox) {
							vx.remove(kx); // 移除以前的记录
							// 添加两条被截断的记录
							vx.put(kx, this.cc.getRowIndex());
							vx.put(this.cc.getRowIndex() + this.cox + (this.cexr * this.cox), vxx + this.cox + (this.cexr * this.cox));
						}
					}
				}
			}
		}
		
	}
	
	private Integer computeExtSize(List<KeyValue<Cell, HashMap<Integer, Integer>>> ml) {
		
		Integer size = 0;
		
		for (KeyValue<Cell, HashMap<Integer, Integer>> kv : ml) {
			
			Integer sx = 0;
			
			Iterator<Integer> i = kv.getValue().keySet().iterator();
			while(i.hasNext()) {
				Integer k = i.next();
				sx += kv.getValue().get(k) - k;
			}
			
			if(sx > size) {
				size = sx;
			}
		}
		
		return size;
	}
	
	private void shiftMap(Sheet sheet, Cell cell, AbstractMap<?, ?> m, Integer depth) {
		// TODO Map 扩展方法
	}
	
	private void setCellValue(Cell c, Object v){
		if(v != null){
			if(v instanceof Integer){
				c.setCellValue((Integer)v);
			}else if(v instanceof Calendar){
				c.setCellValue((Calendar)v);
			}else if(v instanceof Boolean){
				c.setCellValue((Boolean)v);
			}else if(v instanceof Date){
				c.setCellValue((Date)v);
			}else if(v instanceof Double){
				c.setCellValue((Double)v);
			}else if(v instanceof RichTextString){
				c.setCellValue((RichTextString)v);
			}else if(
					v instanceof BigDecimal 
					){
				c.setCellValue(((BigDecimal)v).doubleValue());
			}else if(
					v instanceof BigInteger
					){
				c.setCellValue(((BigInteger)v).longValue());
			}else if(
					v instanceof Long
					) {
				c.setCellValue(((Long)v).longValue());
			}else if(
					v instanceof Short
					) {
				c.setCellValue(((Short)v).shortValue());
			}else if(
					v instanceof Float
					) {
				c.setCellValue(((Float)v).floatValue());
			}else{
				c.setCellValue(v.toString());
			}
		}
		else{
			c.setCellValue("");
		}
	}
	
	private String clearExpressF(String express){
		return express.substring(express.indexOf('{') + 1, express.lastIndexOf('}'));
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private List<Object> executeExpress(String express, Object target){
		
		List<Object> list = new LinkedList<Object>();
		
		if(StringUtils.isNotBlank(express) && target != null){
			if(express.indexOf('.') == -1){
				if(express.indexOf('[') == -1){
					if(target instanceof List){
						for(Object s : (List)target){
							Method get = null;
							try{
								Class clazz = s.getClass();
								get = clazz.getMethod("get", Integer.class);
							}catch(Exception e){}
							if(get != null){
								Object value = null;
								try{
									value = get.invoke(target, Integer.valueOf(express));
								}catch(Exception e){}
								
								if(!(value instanceof List)){
									list.add(value);
								}else{
									list.addAll((List<Object>) list);
								}
							}
						}
					}else{
						Method get = null;
						try{
							Class clazz = target.getClass();
							get = clazz.getMethod("get", Object.class);
						}catch(Exception e){}
						if(get != null){
							Object value = null;
							try{
								value = get.invoke(target, express);
							}catch(Exception e){}
							
							if(value instanceof List){
								list.addAll((List<Object>) value);
							}else{
								list.add(value);
							}
						}
					}
				}else{
					String o = express.substring(0, express.indexOf('['));
					List<Object> listB = new LinkedList<Object>();
					if(StringUtils.isNotBlank(o)){
						listB = executeExpress(o, target);
					}else if(target instanceof List){
						listB = (List<Object>)target;
					}
					if(listB != null && listB.size() > 0){
						String i = express.substring(express.indexOf('[') + 1, express.indexOf(']'));
						Object value = listB.get(Integer.valueOf(i));
						
						if(value instanceof List){
							list.addAll((List<Object>) value);
						}else{
							list.add(value);
						}
						String a = express.substring(express.indexOf(']') + 1);
						// 还有剩余的表达式
						if(StringUtils.isNotBlank(a)){
							executeExpress(a, list);
						}
					}
				}
				
			}else{
				String o = express.substring(0, express.indexOf('.'));
				List<Object> listC = executeExpress(o, target);
				List<Object> listD = null;
				if(listC != null && listC.size() == 1){
					listD = executeExpress(express.substring(express.indexOf('.') + 1), listC.get(0));
				}else if(listC != null && listC.size() > 0){
					listD = executeExpress(express.substring(express.indexOf('.') + 1), listC);
				}
				list.addAll(listD);
			}
		}
		
		return list;
	}
	
	
//	
//	public static String TYPE_FIELD = "field";
//	public static String TYPE_METHOD = "method";
//	public static String TYPE_GET_STRING = "getString";
//	public static String TYPE_ARRAY = "array";
//	
//	public static String TYPE = "type";
//	public static String VALUE = "value";
//	
//	private List<Map<String, Object>> findProccesseLink(String express, Object target){
//		
//		List<Map<String, Object>> link = new LinkedList<Map<String, Object>>();
//		Class clazz = target.getClass();
//		
//		Map<String, Object> tmp = null;
//		
//		if(StringUtils.isNotBlank(express) && target != null){
//			
//			if(express.indexOf(".") == -1){ // 没有小数点
//				if(express.indexOf("[") == -1){ // 没有数组
//					Field field = null;
//					try{
//						field = clazz.getField(express);
//					}catch(Exception e){}
//					
//					if(field != null){ // 找到字段
//						tmp = new HashMap<String, Object>();
//						tmp.put(TYPE, TYPE_FIELD);
//						tmp.put(VALUE, field);
//					}else{
//						Method method = null;
//						try{
//							method = clazz.getMethod(express);
//						}catch(Exception e){}
//						if(method != null){
//							tmp = new HashMap<String, Object>();
//							tmp.put(TYPE, TYPE_FIELD);
//							tmp.put(VALUE, method);
//						}else{
//							Method getM = null;
//							try{
//								getM = clazz.getMethod("get", String.class);
//							}catch(Exception e){}
//							if(getM != null){
//								Object value = null;
//								try{
//									value = getM.invoke(target, express);
//								}catch(Exception e){}
//								if(value != null){
//									
//								}
//							}
//						}
//					}
//						
//					
//				}else{
//					
//				}
//			}else{
//				
//			}
//			
//		}
//		
//		return link;
//	}

	
	
	/**
	 * 获取合并单元格的值
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public String getMergedRegionValue(Sheet sheet, int row, int column) {
		
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {

				if (column >= firstColumn && column <= lastColumn) {
					this.proccessedCas.add(ca); // 将此地址添加到已处理的集合中
					Row fRow = sheet.getRow(firstRow);
					if(fRow != null) {
						Cell fCell = fRow.getCell(firstColumn);
						if(fCell != null) {
							return getCellValue(fCell);
						}
					}
				}
			}
		}
		

		return null;
	}

	/**
	 * 判断合并了行
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	@SuppressWarnings("unused")
	private boolean isMergedRow(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row == firstRow && row == lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}
	private CellRangeAddress getMergedRegionByCell(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return range;
				}
			}
		}
		return null;
	}
	/**
	 * 判断指定的单元格是否是合并单元格
	 * 
	 * @param sheet
	 * @param row
	 *            行下标
	 * @param column
	 *            列下标
	 * @return
	 */
	private boolean isMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}
	/**
	 * 判断指定的单元格是否是合并单元格
	 * 
	 * @param sheet
	 * @param row
	 *            行下标
	 * @param column
	 *            列下标
	 * @return
	 */
	@SuppressWarnings("unused")
	private CellRangeAddress getMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return range;
				}
			}
		}
		return null;
	}
	
	/**
	 * 判断表格是否已经处理过了
	 * @return
	 */
	private Boolean isCellContainsProccessedCas(int row, int column){

		for(CellRangeAddress range : this.proccessedCas){
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
		
	}

	private boolean isExpression(Object value) {
		
		if(
				value instanceof String 
				&& value != null 
				&& ((String)value).startsWith(EXP_START) 
				&& ((String)value).endsWith(EXP_END)
				) {
			
			return true;
			
		}
		
		return false;
	}
	
	/**
	 * 判断sheet页中是否含有合并单元格
	 * 
	 * @param sheet
	 * @return
	 */
	@SuppressWarnings("unused")
	private boolean hasMerged(Sheet sheet) {
		return sheet.getNumMergedRegions() > 0 ? true : false;
	}

	/**
	 * 合并单元格
	 * 
	 * @param sheet
	 * @param firstRow
	 *            开始行
	 * @param lastRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param lastCol
	 *            结束列
	 */
	@SuppressWarnings("unused")
	private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public String getCellValue(Cell cell) {

		if (cell == null)
			return null;

		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

			return cell.getStringCellValue();

		} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

			return String.valueOf(cell.getBooleanCellValue());

		} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

			return cell.getCellFormula();

		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

			return String.valueOf(cell.getNumericCellValue());

		}
		return null;
	}


    /** 
     * 复制一个单元格样式到目的单元格样式 
     * @param fromStyle 
     * @param toStyle 
     */  
    public static void copyCellStyle(CellStyle fromStyle,  
            CellStyle toStyle) {  
    	
    	toStyle.cloneStyleFrom(fromStyle);
//    	
//        toStyle.setAlignment(fromStyle.getAlignment());  
//        //边框和边框颜色  
//        toStyle.setBorderBottom(fromStyle.getBorderBottom());  
//        toStyle.setBorderLeft(fromStyle.getBorderLeft());  
//        toStyle.setBorderRight(fromStyle.getBorderRight());  
//        toStyle.setBorderTop(fromStyle.getBorderTop());  
//        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());  
//        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());  
//        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());  
//        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());  
//          
//        //背景和前景  
//        toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());  
//        toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());  
//          
//        toStyle.setDataFormat(fromStyle.getDataFormat());  
//        toStyle.setFillPattern(fromStyle.getFillPattern());  
////      toStyle.setFont(fromStyle.getFont(null));  
//        toStyle.setHidden(fromStyle.getHidden());  
//        toStyle.setIndention(fromStyle.getIndention());//首行缩进  
//        toStyle.setLocked(fromStyle.getLocked());  
//        toStyle.setRotation(fromStyle.getRotation());//旋转  
//        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());  
//        toStyle.setWrapText(fromStyle.getWrapText());  
          
    }  
    /** 
     * Sheet复制 
     * @param fromSheet 
     * @param toSheet 
     * @param copyValueFlag 
     */  
    public static void copySheet(Workbook wb,Sheet fromSheet, Sheet toSheet,  
            boolean copyValueFlag) {  
        //合并区域处理  
        mergerRegion(fromSheet, toSheet);  
        for (Iterator<Row> rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {  
            Row tmpRow = rowIt.next();  
            Row newRow = toSheet.createRow(tmpRow.getRowNum());  
            //行复制  
            copyRow(wb,tmpRow,newRow,copyValueFlag);  
        }  
    }  
    /** 
     * 行复制功能 
     * @param fromRow 
     * @param toRow 
     */  
    public static void copyRow(Workbook wb,Row fromRow,Row toRow,boolean copyValueFlag){  
        for (Iterator<Cell> cellIt = fromRow.cellIterator(); cellIt.hasNext();) {  
            Cell tmpCell = cellIt.next();  
            Cell newCell = toRow.createCell(tmpCell.getColumnIndex());  
            copyCell(wb,tmpCell, newCell, copyValueFlag);  
        }  
    }  
    /** 
    * 复制原有sheet的合并单元格到新创建的sheet 
    *  
    * @param toSheet 新创建sheet
    * @param fromSheet      原有的sheet
    */  
    public static void mergerRegion(Sheet fromSheet, Sheet toSheet) {  
       int sheetMergerCount = fromSheet.getNumMergedRegions();  
       for (int i = 0; i < sheetMergerCount; i++) {
        CellRangeAddress mergedRegion = fromSheet.getMergedRegion(i);
        toSheet.addMergedRegion(mergedRegion);
       }  
    }  
    
    public static CellStyle createCellStyle(Workbook wb,CellStyle fromStyle){  
    	if(CELL_STYLE_CACHE.containsKey(fromStyle)) {
    		return CELL_STYLE_CACHE.get(fromStyle);
    	}else {
    		CellStyle newstyle = fromStyle;
    		try {
                newstyle = wb.createCellStyle();
                copyCellStyle(fromStyle, newstyle);
    		}catch(Exception e) {
    			System.out.println(e.getMessage());
    		}
            CELL_STYLE_CACHE.put(fromStyle, newstyle);
            return newstyle;
    	}
    }  
    /** 
     * 复制单元格 
     *  
     * @param srcCell 
     * @param distCell 
     * @param copyValueFlag 
     *            true则连同cell的内容一起复制 
     */  
    public static void copyCell(Workbook wb,Cell srcCell, Cell distCell,  
            boolean copyValueFlag) {  
    	if(wb != null && srcCell != null && distCell != null) {

            CellStyle newstyle = createCellStyle(wb, srcCell.getCellStyle());
//            distCell.setEncoding(srcCell.getEncoding());  
            //样式  
            distCell.setCellStyle(newstyle);  
            //评论  
            if (srcCell.getCellComment() != null) {  
                distCell.setCellComment(srcCell.getCellComment());  
            }  
            // 不同数据类型处理  
            int srcCellType = srcCell.getCellType();  
            distCell.setCellType(srcCellType);  
            if (copyValueFlag) {  
                if (srcCellType == Cell.CELL_TYPE_NUMERIC) {  
                    if (DateUtil.isCellDateFormatted(srcCell)) {  
                        distCell.setCellValue(srcCell.getDateCellValue());  
                    } else {  
                        distCell.setCellValue(srcCell.getNumericCellValue());  
                    }  
                } else if (srcCellType == Cell.CELL_TYPE_STRING) {  
                    distCell.setCellValue(srcCell.getRichStringCellValue());  
                } else if (srcCellType == Cell.CELL_TYPE_BLANK) {  
                    // nothing21  
                } else if (srcCellType == Cell.CELL_TYPE_BOOLEAN) {  
                    distCell.setCellValue(srcCell.getBooleanCellValue());  
                } else if (srcCellType == Cell.CELL_TYPE_ERROR) {  
                    distCell.setCellErrorValue(srcCell.getErrorCellValue());  
                } else if (srcCellType == Cell.CELL_TYPE_FORMULA) {  
                    distCell.setCellFormula(srcCell.getCellFormula());  
                } else { // nothing29  
                }  
            }  
    	}
    }  
	
	public Object getModel() {
		return model;
	}

	public void setModel(Object model) {
		this.model = model;
	}

	public String getExcelFilePath() {
		return excelFilePath;
	}

	public void setExcelFilePath(String excelFilePath) {
		this.excelFilePath = excelFilePath;
	}

	public InputStream getInputStream() {
		return inputStream;
	}

	public void setInputStream(InputStream inputStream) {
		this.inputStream = inputStream;
	}

	public Workbook getBook() {
		return book;
	}

	public void setBook(Workbook book) {
		CELL_STYLE_CACHE.clear();
		this.book = book;
	}

	public Set<CellRangeAddress> getProccessedCas() {
		return proccessedCas;
	}

	public void setProccessedCas(Set<CellRangeAddress> proccessedCas) {
		this.proccessedCas = proccessedCas;
	}

	public HSSFWorkbook gethBook() {
		return hBook;
	}

	public Map<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>> getRowRanges() {
		return rowRanges;
	}

	public Map<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>> getSheetsRowRanges() {
		return sheetsRowRanges;
	}

	public void setSheetsRowRanges(
			Map<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>> sheetsRowRanges) {
		this.sheetsRowRanges = sheetsRowRanges;
	}

	public Map<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>> getSheetsCellRanges() {
		return sheetsCellRanges;
	}

	public void setSheetsCellRanges(
			Map<String, HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>>> sheetsCellRanges) {
		this.sheetsCellRanges = sheetsCellRanges;
	}

	public void setRowRanges(HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>> rowRanges) {
		this.rowRanges = rowRanges;
	}

	public void setCellRanges(HashMap<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>> cellRanges) {
		this.cellRanges = cellRanges;
	}

	public Map<Integer, ArrayList<KeyValue<Cell, HashMap<Integer, Integer>>>> getCellRanges() {
		return cellRanges;
	}


	public XSSFWorkbook getxBook() {
		return xBook;
	}

	public void setxBook(XSSFWorkbook xBook) {
		this.xBook = xBook;
	}

	public void sethBook(HSSFWorkbook hBook) {
		this.hBook = hBook;
	}

}
