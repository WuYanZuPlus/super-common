package com.github.wuyanzuplus.excel.core;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

/**
 * @author daniel.hu
 */
@Slf4j
@Component
@SuppressWarnings("squid:MaximumInheritanceDepth")
class ApiXlsxView extends ExcelBaseView {

    public ApiXlsxView() {
        super("接口导出", "接口数据", ApiTemplateEnum.values());
    }

}
