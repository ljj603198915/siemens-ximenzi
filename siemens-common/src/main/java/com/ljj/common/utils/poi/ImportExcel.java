package com.ljj.common.utils.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 导入Excel文件（支持“XLS”和“XLSX”格式）
 *
 * @author lijunjie
 * @version 2012-06-23
 */
public class ImportExcel {

    private static Logger log = LoggerFactory.getLogger(ImportExcel.class);

    /**
     * 工作薄对象
     */
    public Workbook wb;

    /**
     * 工作表对象
     */
	public Sheet sheet;

    /**
     * 标题行号
     */
	public int headerNum;

    /**
     * 构造函数
     *
     * @param fileName  导入文件，读取第一个工作表
     * @param headerNum 标题行号，数据行号=标题行号+1
     * @throws InvalidFormatException
     * @throws IOException
     */
    public ImportExcel(String fileName, int headerNum)
            throws InvalidFormatException, IOException {
        this(new File(fileName), headerNum);
    }

    /**
     * 构造函数
     *
     * @param file      导入文件对象，读取第一个工作表
     * @param headerNum 标题行号，数据行号=标题行号+1
     * @throws InvalidFormatException
     * @throws IOException
     */
    public ImportExcel(File file, int headerNum)
            throws InvalidFormatException, IOException {
        this(file, headerNum, 0);
    }

    /**
     * 构造函数
     *
     * @param fileName   导入文件
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     * @throws InvalidFormatException
     * @throws IOException
     */
    public ImportExcel(String fileName, int headerNum, int sheetIndex)
            throws InvalidFormatException, IOException {
        this(new File(fileName), headerNum, sheetIndex);
    }

    /**
     * 构造函数
     *
     * @param file       导入文件对象
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     * @throws InvalidFormatException
     * @throws IOException
     */
    public ImportExcel(File file, int headerNum, int sheetIndex)
            throws InvalidFormatException, IOException {
        this(file.getName(), new FileInputStream(file), headerNum, sheetIndex);
    }

    /**
     * 构造函数
     *
     * @param fileName   导入文件对象
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     * @throws InvalidFormatException
     * @throws IOException
     */
    public ImportExcel(String fileName, InputStream is, int headerNum, int sheetIndex)
            throws InvalidFormatException, IOException {
        if (StringUtils.isBlank(fileName)) {
            throw new RuntimeException("导入文档为空!");
        } else if (fileName.toLowerCase().endsWith("xls")) {
            this.wb = new HSSFWorkbook(is);
        } else if (fileName.toLowerCase().endsWith("xlsx")) {
            this.wb = new XSSFWorkbook(is);
        } else {
            throw new RuntimeException("文档格式不正确!");
        }
        if (this.wb.getNumberOfSheets() < sheetIndex) {
            throw new RuntimeException("文档中没有工作表!");
        }
        this.sheet = this.wb.getSheetAt(sheetIndex);
        this.headerNum = headerNum;
        log.debug("Initialize success.");
    }

    /**
     * 获取行对象
     *
     * @param rownum
     * @return
     */
    public Row getRow(int rownum) {
        return this.sheet.getRow(rownum);
    }

    /**
     * 获取数据行号
     *
     * @return
     */
    public int getDataRowNum() {
        return headerNum + 1;
    }

    /**
     * 获取最后一个数据行号
     *
     * @return
     */
    public int getLastDataRowNum() {
        return this.sheet.getLastRowNum() + headerNum;
    }

    /**
     * 获取最后一个列号
     *
     * @return
     */
    public int getLastCellNum() {
        return this.getRow(headerNum).getLastCellNum();
    }

    public String getCellValue(Row row, int column) {
        try {
            Cell cell = row.getCell(column);
            return getCellValue(cell);
        } catch (Exception e) {
            System.out.println(e.getMessage());
            return null;
        }
    }

    public String getCellValue(Cell cell) {
        if (cell == null) return null;
        String val = null;
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case _NONE:
                break;
            case BLANK:
                val = "";
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //用于转化为日期格式
                    Date d = cell.getDateCellValue();
                    DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                    val = formater.format(d);
                } else {
                    val = NumberToTextConverter.toText(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                val = cell.getCellFormula();
                break;
            case STRING:
                val = cell.getStringCellValue();
                break;
            default:
                val = cell.getStringCellValue();
        }

        return val;
    }
}
