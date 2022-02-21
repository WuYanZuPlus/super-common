package com.github.wuyanzuplus.excel.core;

import lombok.AllArgsConstructor;

/**
 * ExcelImportErrorEnum
 *
 * @author daniel.hu
 */
@AllArgsConstructor
public enum ExcelImportErrorEnum {
    FILE_TITLE_ERROR("文件表头错误，请按模板重新上传"),
    NULL_VALUE("空值"),
    OVER_LENGTH("超过长度限制"),
    DATE_FORMAT_ERROR("日期格式错误"),
    CONTENT_BEYOND_RANGE("内容超出选项范围"),
    ;

    String value;
}
