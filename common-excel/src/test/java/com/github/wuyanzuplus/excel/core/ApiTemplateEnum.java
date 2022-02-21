package com.github.wuyanzuplus.excel.core;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 接口导入导出枚举
 *
 * @author daniel.hu
 */
@Getter
@AllArgsConstructor
public enum ApiTemplateEnum implements ExcelHandler {
    PROJECT("项目名", "project", true, 50),
    API_NAME("接口名称", "apiName", true, 30),
    API_CODE("接口编码", "apiCode", true, 100),
    API_URL("接口地址", "apiUrl", true, 500),
    API_PLATFORM("资源属性", "apiPlatform", true, 10, val -> {
        try {
            Platform.valueOf(val);
        } catch (IllegalArgumentException e) {
            return ExcelImportErrorEnum.CONTENT_BEYOND_RANGE.name();
        }
        return null;
    }, (ValueImportResolver<Platform>) val -> {
        try {
            return Platform.valueOf(val);
        } catch (IllegalArgumentException e) {
            return null;
        }
    }, val -> val == null ? "" : val.toString());

    /**
     * 标题名称
     */
    private String titleName;
    /**
     * 字段名称
     */
    private String fieldName;
    /**
     * 是否必填
     */
    private boolean required;
    /**
     * 最大长度
     */
    private int maxLength;

    /**
     * 校验（便于错误统计）
     */
    private ValueValidator valueValidator;
    /**
     * 导入时，对value的处理
     */
    private ValueImportResolver valueImportResolver;
    /**
     * 导出时，对value的处理
     */
    private ValueExportResolver valueExportResolver;

    ApiTemplateEnum(String titleName, String fieldName, boolean required, int maxLength) {
        this.titleName = titleName;
        this.fieldName = fieldName;
        this.required = required;
        this.maxLength = maxLength;
    }

}
