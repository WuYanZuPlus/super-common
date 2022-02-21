package com.github.wuyanzuplus.excel.core;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mockito;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assertions.*;


/**
 * @author daniel.hu
 */
@Slf4j
public class ExcelImportAndExportApiTest {

    private ApiXlsxView view;

    @Before
    public void setUp() {
        view = new ApiXlsxView();
    }

    private static List<ApiEntity> getInitApiExportList() {
        return Arrays.asList(
                newApi(Platform.系统, "api1", "test_add", true, "权限管理", "/abc/def1"),
                newApi(Platform.系统, "api2", "test_update", true, "权限管理", "/abc/def2"),
                newApi(Platform.租户, "api3", "test_delete", true, "会员", "/abc/def3")
        );
    }

    private static ApiEntity newApi(Platform apiPlatform, String apiName, String apiCode, Boolean valid, String project, String apiUrl) {
        ApiEntity api = new ApiEntity();
        api.setApiPlatform(apiPlatform);
        api.setApiName(apiName);
        api.setApiCode(apiCode);
        api.setValid(valid);
        api.setProject(project);
        api.setApiUrl(apiUrl);
        return api;
    }

    /**
     * 导入功能
     */
    @SneakyThrows
    private List<ApiEntity> importApi(MultipartFile file) {
        List<String[]> list = ExcelUtil.readExcelWithFirstSheet(file);
        List<ApiEntity> entities = new ArrayList<>();
        if (!list.isEmpty()) {
            if (!ExcelUtil.isTitleLegal(list.get(0), ApiTemplateEnum.values())) {
                throw new ExcelResolvingException(ExcelImportErrorEnum.FILE_TITLE_ERROR.value);
            }
            for (int i = 1; i < list.size(); i++) {
                String[] rowData = list.get(i);
                if (!ExcelUtil.isRowLegal(rowData, ApiTemplateEnum.values())) {
                    continue;
                }
                ApiEntity apiEntity = new ApiEntity();
                entities.add(ExcelUtil.transformData(apiEntity, rowData, ApiTemplateEnum.values()));
            }
        }
        return entities;
    }

    @Test
    public void 导入_正常数据_success() throws IOException {

        MockMultipartFile file = new MockMultipartFile("file", "资源导入模板_正常.xlsx", "multipart/form-data", Object.class.getResourceAsStream("/资源导入模板_正常.xlsx"));
        List<ApiEntity> entities = importApi(file);

        assertEquals(2, entities.size());
        assertThat(entities.get(0))
                .hasFieldOrPropertyWithValue("project", "权限管理")
                .hasFieldOrPropertyWithValue("apiName", "接口1")
                .hasFieldOrPropertyWithValue("apiCode", "code1")
                .hasFieldOrPropertyWithValue("apiUrl", "post:/project/user/api1");

        assertThat(entities.get(1))
                .hasFieldOrPropertyWithValue("project", "权限管理")
                .hasFieldOrPropertyWithValue("apiName", "接口2")
                .hasFieldOrPropertyWithValue("apiCode", "code2")
                .hasFieldOrPropertyWithValue("apiUrl", "post:/project/user/api2");

    }

    @Test
    public void test_export_success() throws Exception {
        Map<String, Object> map = new HashMap<>();
        map.put("data", getInitApiExportList());
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        HttpServletRequest request = Mockito.mock(HttpServletRequest.class);
        HttpServletResponse response = Mockito.mock(HttpServletResponse.class);
        Mockito.when(request.getHeader("User-Agent"))
                .thenReturn("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36");

        view.buildExcelDocument(map, workbook, request, response);

        // sheet1
        SXSSFSheet sheet1 = workbook.getSheetAt(0);
        assertThat(sheet1.getSheetName()).isEqualTo("接口数据");

        // row1 -->标题行
        SXSSFRow row1 = sheet1.getRow(0);
        assertThat(row1.getCell(0).toString()).isEqualTo(ApiTemplateEnum.PROJECT.getTitleName());
        assertThat(row1.getCell(1).toString()).isEqualTo(ApiTemplateEnum.API_NAME.getTitleName());
        assertThat(row1.getCell(2).toString()).isEqualTo(ApiTemplateEnum.API_CODE.getTitleName());
        assertThat(row1.getCell(3).toString()).isEqualTo(ApiTemplateEnum.API_URL.getTitleName());
        assertThat(row1.getCell(4).toString()).isEqualTo(ApiTemplateEnum.API_PLATFORM.getTitleName());

        // row2 -->content
        SXSSFRow row2 = sheet1.getRow(1);
        assertThat(row2.getCell(0).toString()).isEqualTo("权限管理");
        assertThat(row2.getCell(1).toString()).isEqualTo("api1");
        assertThat(row2.getCell(2).toString()).isEqualTo("test_add");
        assertThat(row2.getCell(3).toString()).isEqualTo("/abc/def1");
        assertThat(row2.getCell(4).toString()).isEqualTo("系统");
    }

    @Test
    public void 导入_空数据_error() throws IOException {
        MockMultipartFile file = new MockMultipartFile("file", "资源导入模板_空数据.xlsx", "multipart/form-data", Object.class.getResourceAsStream("/资源导入模板_空数据.xlsx"));
        List<ApiEntity> entities = importApi(file);
        assertTrue(entities.isEmpty());
    }

    @Test
    public void 导入_标题错误_error() throws IOException {
        MockMultipartFile file = new MockMultipartFile("file", "资源导入模板_标题错误.xlsx", "multipart/form-data", Object.class.getResourceAsStream("/资源导入模板_标题错误.xlsx"));
        try {
            importApi(file);
            fail();
        } catch (ExcelResolvingException e) {
            assertEquals(ExcelImportErrorEnum.FILE_TITLE_ERROR.value, e.getMessage());
        }
    }

    @Test
    public void 导入_内容异常_error() throws IOException {
        // 内容异常不抛出,可记录
        MockMultipartFile file = new MockMultipartFile("file", "资源导入模板_内容异常.xlsx", "multipart/form-data", Object.class.getResourceAsStream("/资源导入模板_内容异常.xlsx"));
        List<ApiEntity> entities = importApi(file);

        assertEquals(1, entities.size());
        assertThat(entities.get(0))
                .hasFieldOrPropertyWithValue("project", "权限管理")
                .hasFieldOrPropertyWithValue("apiName", "接口2")
                .hasFieldOrPropertyWithValue("apiCode", "code2")
                .hasFieldOrPropertyWithValue("apiUrl", "post:/project/user/api2");
    }
}
