package com.github.wuyanzuplus.excel.core;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * @Author: daniel.hu
 * @Date: 2020/9/4 14:06
 */
@Slf4j
@SuppressWarnings({"unused"})
public class ExcelUtil {
    /**
     * excel导出数据key
     */
    public static final String EXCEL_EXPORT_DATA_KEY = "data";

    private ExcelUtil() {
    }

    /**
     * 读取Excel首个sheet页
     */
    @SneakyThrows
    public static List<String[]> readExcelWithFirstSheet(MultipartFile file) {
        return readExcelWithFirstSheet(file.getOriginalFilename(), file.getInputStream());
    }

    /**
     * 读取Excel首个sheet页
     */
    public static List<String[]> readExcelWithFirstSheet(String filename, InputStream inputStream) {
        Workbook workBook = createWorkBook(filename, inputStream);
        return readExcelWithFirstSheet(workBook);
    }

    /**
     * 读取Excel所有sheet页
     */
    @SneakyThrows
    public static Map<String, List<String[]>> readExcelWithAllSheet(MultipartFile file) {
        Workbook workBook = createWorkBook(file.getOriginalFilename(), file.getInputStream());
        return readExcelWithAllSheet(workBook);
    }

    /**
     * 创建workbook进行excel读取操作
     */
    @SneakyThrows
    private static Workbook createWorkBook(String filename, InputStream is) {
        return "xls".equalsIgnoreCase(getSuffix(filename)) ? new HSSFWorkbook(is) : new XSSFWorkbook(is);
    }

    /**
     * 获取文件名后缀
     */
    private static String getSuffix(String filename) {
        String[] parts = StringUtils.split(filename, ".");
        return parts.length > 0 ? parts[parts.length - 1] : "";
    }

    /**
     * 读取excel第一个sheet内容
     */
    private static List<String[]> readExcelWithFirstSheet(Workbook workBook) {
        Sheet sheet0 = workBook.getSheetAt(0);
        List<String[]> list = new LinkedList<>();
        if (sheet0 == null || sheet0.getLastRowNum() <= 0) {
            return list;
        }
        resolveSingleSheet(sheet0, list);
        return list;
    }

    /**
     * 读取excel所有sheet内容
     */
    private static Map<String, List<String[]>> readExcelWithAllSheet(Workbook workBook) {
        Map<String, List<String[]>> sheetMap = new HashMap<>();
        int sheetNum = workBook.getNumberOfSheets();
        for (int index = 0; index < sheetNum; index++) {
            Sheet sheet = workBook.getSheetAt(index);
            if (sheet == null || sheet.getLastRowNum() <= 0) {
                continue;
            }
            List<String[]> list = new LinkedList<>();
            resolveSingleSheet(sheet, list);
            sheetMap.put(sheet.getSheetName(), list);
        }
        return sheetMap;
    }

    /**
     * 解析单个sheet
     */
    private static void resolveSingleSheet(Sheet sheet, List<String[]> list) {
        short cellNum = sheet.getRow(0).getLastCellNum();
        for (Row row : sheet) {
            List<String> rowData = new LinkedList<>();
            for (int i = 0; i < cellNum; i++) {
                rowData.add(formatCell(row.getCell(i)));
            }
            list.add(rowData.toArray(new String[0]));
        }
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

    /**
     * 判断标题是否合法
     *
     * @param titles 标题行数组（只针对单行标题）
     * @param metas  枚举值（all）
     * @return true legal
     */
    public static boolean isTitleLegal(String[] titles, ExcelHandler[] metas) {
        for (int i = 0, length = metas.length; i < length; i++) {
            if (!metas[i].getTitleName().equals(titles[i])) {
                return false;
            }
        }
        return true;
    }

    /**
     * 判断excel行数据是否合法
     *
     * @param rowData 行数据
     * @param values  枚举值（all）
     * @return true legal
     */
    public static boolean isRowLegal(String[] rowData, ExcelHandler[] values) {
        for (int i = 0; i < values.length; i++) {
            ExcelHandler columnEnum = values[i];
            String str = rowData[i];
            String error = columnEnum.checkValue(str);
            if (StringUtils.isNotBlank(error)) {
                return false;
            }
        }
        return true;
    }

    /**
     * 导入: 将excel行数据转换为对应实体属性值 (只考虑当前类)
     *
     * @param target  目标实体
     * @param rowData excel行数据
     * @param values  枚举值（all）
     * @return 目标实体
     */
    public static <T> T transformDeclaredData(T target, String[] rowData, ExcelHandler[] values) {
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

    /**
     * 导入: 将excel行数据转换为对应实体属性值 (考虑父类)
     *
     * @param target  目标实体
     * @param rowData excel行数据
     * @param values  枚举值（all）
     * @return 目标实体
     */
    public static <T> T transformData(T target, String[] rowData, ExcelHandler[] values) {
        for (int i = 0; i < values.length; i++) {
            String datum = rowData[i];
            ExcelHandler value = values[i];
            try {
                FieldUtils.writeField(target, value.getFieldName(), value.resolveImportValue(datum), true);
            } catch (IllegalAccessException e) {
                log.error("安全权限异常", e);
            }
        }
        return target;
    }
}
