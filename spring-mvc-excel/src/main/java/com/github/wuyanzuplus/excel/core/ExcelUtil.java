package com.github.wuyanzuplus.excel.core;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

/**
 * @author daniel.hu
 */
@Slf4j
public class ExcelUtil {

    public static List<String[]> readExcel(String filename, InputStream inputStream) {
        Workbook workBook = createWorkBook(filename, inputStream);
        return readExcel(workBook);
    }

    @SneakyThrows
    private static Workbook createWorkBook(String filename, InputStream is) {
        return "xls".equalsIgnoreCase(resolveSuffix(filename)) ? new HSSFWorkbook(is) : new XSSFWorkbook(is);
    }

    private static String resolveSuffix(String filename) {
        String[] parts = StringUtils.split(filename, ".");
        return parts.length > 0 ? parts[parts.length - 1] : "";
    }

    private static List<String[]> readExcel(Workbook workBook) {
        List<String[]> list = new LinkedList<>();
        Sheet sheet0 = workBook.getSheetAt(0);
        if (sheet0 == null || sheet0.getLastRowNum() <= 0) {
            return list;
        }
        short cellNum = sheet0.getRow(0).getLastCellNum();
        for (Row row : sheet0) {
            List<String> rowData = new LinkedList<>();
            for (int i = 0; i < cellNum; i++) {
                rowData.add(formatCell(row.getCell(i)));
            }
            list.add(rowData.toArray(new String[0]));
        }
        return list;
    }

    /**
     * 处理单元格格式
     *
     * @param cell 单元格
     * @return string
     */
    private static String formatCell(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            // 数值类型(whole numbers, fractional numbers, dates)
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return DateFormatUtils.format(DateUtil.getJavaDate(cell.getNumericCellValue()), "yyyy-MM-dd");
                } else {
                    cell.setCellType(CellType.STRING);
                    String temp = cell.getStringCellValue();
                    // 判断是否包含小数点，如果不含小数点，则以字符串读取，如果含小数点，则转换为Double类型的字符串
                    if (temp.contains(".")) {
                        return String.valueOf(new Double(temp)).trim();
                    } else {
                        return temp.trim();
                    }
                }
            case STRING:
                return cell.getStringCellValue();
            // 表达式
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            case BOOLEAN:
                return cell.getBooleanCellValue() + "";
            case ERROR:
                return cell.getErrorCellValue() + "";
            default:
        }
        return "";
    }

    public static <T> T transformData(T target, String[] rowData, ExcelHandler[] values) {
        for (int i = 0; i < values.length; i++) {
            String datum = rowData[i];
            ExcelHandler value = values[i];
            try {
                FieldUtils.writeDeclaredField(target, value.getFieldName(), value.resolveImportValue(datum), true);
            } catch (IllegalAccessException e) {
                log.error("安全权限异常", e);
            }
        }
        return target;
    }

    public static boolean checkTitle(String[] titles, ExcelHandler[] metas) {
        for (int i = 0, length = metas.length; i < length; i++) {
            if (!metas[i].getTitleName().equals(titles[i])) {
                return false;
            }
        }
        return true;
    }
}
