package com.baobeidaodao.poi.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author DaoDao
 */
@Slf4j
public class ExcelUtil {

    public final static String XLS = "xls";
    public final static String XLSX = "xlsx";

    private static final SimpleDateFormat SIMPLE_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    private static void checkExcel(File file) {
        FileUtil.checkFile(file);
        String fileName = file.getName();
        if (!fileName.endsWith(XLS) && !fileName.endsWith(XLSX)) {
            log.error(fileName + " is not excel file!");
            try {
                throw new IOException(fileName + " is not excel file!");
            } catch (IOException e) {
                e.printStackTrace();
                log.error(e.getMessage());
            }
        }
    }

    private static Workbook readFile(File file) {
        checkExcel(file);
        String fileName = file.getName();
        Workbook workbook = null;
        try {
            InputStream inputStream = new FileInputStream(file);
            if (fileName.endsWith(XLSX)) {
                workbook = new XSSFWorkbook(inputStream);
            } else if (fileName.endsWith(XLS)) {
                workbook = new HSSFWorkbook(inputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
            log.error(e.getMessage());
        }
        return workbook;
    }

    private static void writeFile(Workbook workbook, File file) {
        try {
            OutputStream fileOut = new FileOutputStream(file.getPath());
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
            log.error(e.getMessage());
        }
    }

    public static Map<String, List<Object[]>> readExcel(File file) {
        // return readExcelWithoutBlank(file);
        return readExcelWithBlank(file);
    }

    private static Map<String, List<Object[]>> readExcelWithBlank(File file) {
        Workbook workbook = readFile(file);
        if (null == workbook) {
            return null;
        }
        Map<String, List<Object[]>> data = new HashMap<>(1 << 4);
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            if (null == sheet) {
                data.put(null, null);
                continue;
            }
            String sheetName = sheet.getSheetName();
            List<Object[]> rowList = new ArrayList<>();
            int lastRowNum = sheet.getLastRowNum();
            for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (null == row) {
                    rowList.add(new String[0]);
                    continue;
                }
                Object[] valueArray = new Object[lastRowNum];
                int lastCellNum = row.getLastCellNum();
                for (int cellIndex = 0; cellIndex <= lastCellNum; cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    if (null == cell) {
                        valueArray[cellIndex] = null;
                        continue;
                    }
                    Object value = readCell(cell);
                    valueArray[cellIndex] = value;
                }
                rowList.add(valueArray);
            }
            data.put(sheetName, rowList);
        }
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
            log.error(e.getMessage());
        }
        return data;
    }

    private static Object readCell(Cell cell) {
        if (null == cell) {
            return null;
        }
        Object value;
        switch (cell.getCellType()) {
            case _NONE:
                value = null;
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case BLANK:
                value = "";
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case ERROR:
                value = cell.getErrorCellValue();
                break;
            default:
                value = null;
                break;
        }
        return value;
    }

    public static File writeExcel(Map<String, List<Object[]>> data, String filePath) {
        File file = new File(filePath);
        checkExcel(file);
        if (null == data) {
            return file;
        }
        // Workbook workbook = new HSSFWorkbook();
        Workbook workbook = new XSSFWorkbook();
        /* CreationHelper createHelper = workbook.getCreationHelper(); */
        for (Map.Entry<String, List<Object[]>> entry : data.entrySet()) {
            String sheetName = entry.getKey();
            List<Object[]> rowList = entry.getValue();
            Sheet sheet = workbook.createSheet(sheetName);
            int size = rowList.size();
            for (int rowIndex = 0; rowIndex < size; rowIndex++) {
                Object[] valueArray = rowList.get(rowIndex);
                Row row = sheet.createRow(rowIndex);
                int length = valueArray.length;
                for (int cellIndex = 0; cellIndex < length; cellIndex++) {
                    Cell cell = row.createCell(cellIndex);
                    Object value = valueArray[cellIndex];
                    writeCell(cell, value);
                }
            }
        }
        writeFile(workbook, file);
        return file;
    }


    private static void writeCell(Cell cell, Object value) {
        if (null == value || "".equals(value)) {
            /// cell.setCellType(CellType._NONE);
            cell.setCellType(CellType.BLANK);
        } else if (value instanceof Number) {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(Double.valueOf(String.valueOf(value)));
        } else if (value instanceof String) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(String.valueOf(value));
///        } else if (value instanceof String) {
///            cell.setCellType(CellType.FORMULA);
        } else if (value instanceof Boolean) {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Error || value instanceof Exception) {
            cell.setCellType(CellType.ERROR);
        } else if (value instanceof Date) {
            cell.setCellType(CellType.STRING);
            Date date = (Date) value;
            String dataTimeString;
            dataTimeString = SIMPLE_DATE_FORMAT.format(date);
            cell.setCellValue(dataTimeString);
        } else if (value instanceof Calendar) {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue((Calendar) value);
        } else {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(String.valueOf(value));
        }
    }

}
