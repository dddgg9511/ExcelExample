package github.choo.excel;

import github.choo.excel.data.ExcelData;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

public class ExcelXlsxView extends AbstractXlsxView {
    private final ExcelData excelData;

    public ExcelXlsxView(ExcelData excelData) {
        this.excelData = excelData;
    }

    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
        new ExcelWriter(workbook, excelData, response).create();
    }
}
