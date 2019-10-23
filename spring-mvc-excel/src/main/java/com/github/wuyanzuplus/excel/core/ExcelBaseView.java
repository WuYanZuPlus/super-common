package com.github.wuyanzuplus.excel.core;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpHeaders;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * @author daniel.hu
 */
@Slf4j
@SuppressWarnings("squid:MaximumInheritanceDepth")
@NoArgsConstructor
@AllArgsConstructor
public abstract class ExcelBaseView<T> extends AbstractXlsxView {

    private String filename;

    private String sheetname;

    private ExcelHandler[] excelHandlers = new ExcelHandler[0];

    @Override
    protected void buildExcelDocument(Map<String, Object> map, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        String filename0 = getFilename(map) + ".xlsx";
        // name.getBytes("UTF-8")处理safari的乱码问题
        String userAgent = request.getHeader(HttpHeaders.USER_AGENT);
        byte[] bytes = userAgent.contains("MSIE") ? filename0.getBytes() : filename0.getBytes(StandardCharsets.UTF_8);
        filename0 = new String(bytes, StandardCharsets.ISO_8859_1);
        // 文件名外的双引号处理firefox的空格截断问题
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, String.format("attachment; filename=\"%s\"", filename0));

        Sheet sheet = workbook.createSheet(getSheetname(map));
        sheet.setDefaultColumnWidth(30);

        AtomicInteger rowCount = new AtomicInteger(0);
        Row title = sheet.createRow(rowCount.getAndIncrement());
        setTitle(title);

        @SuppressWarnings("unchecked")
        List<T> data = (List<T>) map.get("data");
        for (T datum : data) {
            Row row = sheet.createRow(rowCount.getAndIncrement());
            setRow(row, datum);
        }
    }

    /**
     * 设置标题
     */
    protected void setTitle(Row row) {
        ExcelHandler[] metas = getExcelHandlers();
        for (int i = 0; i < metas.length; i++) {
            row.createCell(i).setCellValue(metas[i].getTitleName());
        }
    }

    /**
     * 设置行数据
     */
    protected void setRow(Row row, T datum) {
        ExcelHandler[] metas = getExcelHandlers();
        for (int i = 0; i < metas.length; i++) {
            ExcelHandler excelHandler = metas[i];
            Object val = readField(datum, excelHandler);
            row.createCell(i).setCellValue(excelHandler.resolveExportValue(val));
        }
    }

    protected Object readField(T datum, ExcelHandler excelHandler) {
        try {
            return FieldUtils.readDeclaredField(datum, excelHandler.getFieldName(), true);
        } catch (IllegalAccessException e) {
            log.warn("字段导出失败: {} - {}", excelHandler.getTitleName(), e.getMessage());
        }
        return null;
    }

    protected String getFilename(Map<String, Object> map) {
        return (String) map.getOrDefault("filename", filename);
    }

    protected String getSheetname(Map<String, Object> map) {
        return (String) map.getOrDefault("sheetname", sheetname);
    }

    protected ExcelHandler[] getExcelHandlers() {
        return excelHandlers;
    }

}


