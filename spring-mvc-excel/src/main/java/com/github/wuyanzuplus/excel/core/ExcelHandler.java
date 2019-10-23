package com.github.wuyanzuplus.excel.core;

import org.apache.commons.lang3.StringUtils;

/**
 * @author daniel.hu
 */
public interface ExcelHandler {
    String getTitleName();

    String getFieldName();

    default boolean isRequired() {
        return false;
    }

    int getMaxLength();

    default ValueValidator getValueValidator() {
        return null;
    }

    default ValueImportResolver getValueImportResolver() {
        return null;
    }

    default ValueExportResolver getValueExportResolver() {
        return null;
    }

    /**
     * 值校验
     */
    default String checkValue(String val) {
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
        } else {
            return val == null ? "" : val.toString();
        }
    }
}
