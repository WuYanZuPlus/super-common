package com.github.wuyanzuplus.excel.core;

import org.apache.commons.lang3.StringUtils;

/**
 * @author daniel.hu
 */
public interface ExcelHandler {
    /**
     * 获取Excel首行标题名称
     */
    String getTitleName();

    /**
     * 获取Excel标题对应的属性名
     */
    String getFieldName();

    /**
     * 是否必填（默认非必填）
     */
    default boolean isRequired() {
        return false;
    }

    /**
     * 字段最大长度
     */
    int getMaxLength();

    /**
     * 获取校验器
     */
    default ValueValidator getValueValidator() {
        return null;
    }

    /**
     * 获取导入值处理器
     */
    default ValueImportResolver getValueImportResolver() {
        return null;
    }

    /**
     * 获取导出值处理器
     */
    default ValueExportResolver getValueExportResolver() {
        return null;
    }

    /**
     * 导入时值校验
     */
    default String checkImportValue(String val) {
        if (isRequired() && StringUtils.isBlank(val)) {
            return ExcelImportErrorEnum.NULL_VALUE.name();
        }
        if (StringUtils.length(val) > getMaxLength()) {
            return ExcelImportErrorEnum.OVER_LENGTH.name();
        }
        if (StringUtils.isNotBlank(val)) {
            ValueValidator valueValidator = getValueValidator();
            if (valueValidator != null) {
                return valueValidator.checkValue(val);
            }
        }
        return null;
    }

    /**
     * 导入时值处理
     */
    default Object resolveImportValue(String val) {
        ValueImportResolver valueImportResolver = getValueImportResolver();
        return valueImportResolver != null ? valueImportResolver.importResolve(val) : val;
    }

    /**
     * 导出时值处理
     */
    default String resolveExportValue(Object val) {
        ValueExportResolver valueExportResolver = getValueExportResolver();
        if (valueExportResolver != null) {
            return valueExportResolver.exportResolve(val);
        }
        return val == null ? "" : val.toString();
    }
}
