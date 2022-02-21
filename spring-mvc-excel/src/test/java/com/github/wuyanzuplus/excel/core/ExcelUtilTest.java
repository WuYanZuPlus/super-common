package com.github.wuyanzuplus.excel.core;

import org.apache.poi.openxml4j.opc.ContentTypes;
import org.apache.poi.openxml4j.opc.internal.ContentType;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.io.IOException;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.*;

/**
 * @author daniel.hu
 */
public class ExcelUtilTest {

    @Test
    public void 处理单元格格式() throws IOException {
        MockMultipartFile file = new MockMultipartFile("file", "资源导入模板_格式转换.xlsx", "multipart/form-data", Object.class.getResourceAsStream("/资源导入模板_格式转换.xlsx"));
        List<String[]> list = ExcelUtil.readExcelWithFirstSheet(file.getOriginalFilename(), file.getInputStream());

        assertEquals(2, list.size());
        String[] array1 = list.get(0);
        assertThat(array1.length).isEqualTo(6);
        assertThat(array1[0]).isEqualTo("项目名");
        assertThat(array1[1]).isEqualTo("接口名称");
        assertThat(array1[2]).isEqualTo("接口编码");
        assertThat(array1[3]).isEqualTo("人民币");
        assertThat(array1[4]).isEqualTo("资源属性");
        assertThat(array1[5]).isEqualTo("日期");

        String[] array2 = list.get(1);
        assertThat(array2.length).isEqualTo(6);
        assertThat(array2[0]).isEqualTo("权限管理");
        assertThat(array2[1]).isEqualTo("接口1");
        assertThat(array2[2]).isEqualTo("");
        assertThat(array2[3]).isEqualTo("100");
        assertThat(array2[4]).isEqualTo("租户");
        assertThat(array2[5]).isEqualTo("2018-01-01");

    }

}