package com.ljj.common.utils.poi;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

public class ExcelUtilLh {

//    private static final Logger log = LoggerFactory.getLogger(ExcelUtilLh.class);
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    /**
     * 判断Excel的版本,获取Workbook
     *
     * @param pathName
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbok(String pathName) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(pathName);
        Workbook wb = null;
        if (pathName.endsWith(EXCEL_XLS)) { // Excel 2003
            wb = new HSSFWorkbook(fileInputStream);
        } else if (pathName.endsWith(EXCEL_XLSX)) { // Excel 2007/2010
            wb = new XSSFWorkbook(fileInputStream);
        }
        return wb;
    }

    /**
     * 判断Excel的版本,获取Workbook
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbok(File file)
            throws IOException {
        String name = file.getName();
        Workbook wb = null;
        if (name.endsWith(EXCEL_XLS)) { // Excel 2003
            wb = new HSSFWorkbook(new FileInputStream(file));
        } else if (name.endsWith(EXCEL_XLSX)) { // Excel 2007/2010
            wb = new XSSFWorkbook(new FileInputStream(file));
        }
        return wb;
    }

    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell
     * @return
     */
    public static String getStringCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellType = cell.getCellType();
        String cellValue = null;
        switch (cellType) {
            case _NONE:
                cellValue = null;
                break;
            case BLANK:
                cellValue = "";
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    //用于转化为日期格式
                    Date d = cell.getDateCellValue();
                    DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = formater.format(d);
                    break;
                } else {
                    cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                    break;
                }
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            default:
                cellValue = cell.getStringCellValue();
        }
        return cellValue.trim();
    }

    /**
     * 绿色信贷和环境效益对公式特殊处理
     *
     * @param cell
     * @return
     */
    public static String getStringTsCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellType = cell.getCellType();
        String cellValue = null;
        switch (cellType) {
            case _NONE:
                cellValue = null;
                break;
            case BLANK:
                cellValue = "";
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    //用于转化为日期格式
                    Date d = cell.getDateCellValue();
                    DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = formater.format(d);
                    break;
                } else {
                    cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                    break;
                }
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case FORMULA:
                try {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            default:
                cellValue = cell.getStringCellValue();
        }
        return cellValue.trim();
    }

    /**
     * 编码文件名
     */
    public static String encodingFilename(String filename) {
        filename = UUID.randomUUID().toString() + "_" + filename + ".xlsx";
        return filename;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     *
     * @param sheet
     * @param row    行下标
     * @param column 列下标
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
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
     * 判断单元格是否为合并单元格，是的话则将单元格的值返回
     *
     * @param listCombineCell 存放合并单元格的list
     * @param cell            需要判断的单元格
     * @param sheet           sheet
     * @return
     */
    public static String isCombineCell(List<CellRangeAddress> listCombineCell, Cell cell, Sheet sheet) {
        int firstC = 0;
        int lastC = 0;
        int firstR = 0;
        int lastR = 0;
        String cellValue = null;
        for (CellRangeAddress ca : listCombineCell) {
            //获得合并单元格的起始行, 结束行, 起始列, 结束列
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            if (cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR) {
                if (cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC) {
                    Row fRow = sheet.getRow(firstR);
                    Cell fCell = fRow.getCell(firstC);
                    cellValue = getStringCellValue(fCell);
                    break;
                }
            } else {
                cellValue = "";
            }
        }
        return cellValue;
    }

    /**
     * 合并单元格处理,获取合并行
     *
     * @param sheet
     * @return List<CellRangeAddress>
     */
    public static List<CellRangeAddress> getCombineCell(Sheet sheet) {
        List<CellRangeAddress> list = new ArrayList();
        //获得一个 sheet 中合并单元格的数量
        int sheetmergerCount = sheet.getNumMergedRegions();
        //遍历所有的合并单元格
        for (int i = 0; i < sheetmergerCount; i++) {
            //获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }
}
