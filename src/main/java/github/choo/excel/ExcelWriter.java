package github.choo.excel;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import github.choo.excel.annotation.ExcelColumnName;
import github.choo.excel.data.ExcelData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelWriter {
    private final Workbook workbook;
    private final ExcelData data;
    private final HttpServletResponse response;
    private static ObjectMapper objectMapper;

    static {
        objectMapper = new ObjectMapper();
        objectMapper.registerModule(new JavaTimeModule());

    }

    public ExcelWriter(Workbook workbook,ExcelData data, HttpServletResponse response) {
        this.workbook = workbook;
        this.data = data;
        this.response = response;
    }

    public void create(){
        setFileName(response, mapToFileName());
        Sheet sheet = workbook.createSheet();

        createHead(sheet, mapToHeadList());
        createBody(sheet, mapToBodyList());
    }

    private String mapToFileName(){
        return String.valueOf(data.getFileName());
    }

    private List<String> mapToHeadList(){
        return data.getHeaders();
    }

    private List<List<String>> mapToBodyList(){
        return data.getBodies();
    }

    private void setFileName(HttpServletResponse response, String fileName) {
        response.setHeader("Content-Disposition",
                "attachment; filename=\"" + getFileExtension(fileName) + "\"");
    }

    private String getFileExtension(String fileName) {
        if (workbook instanceof XSSFWorkbook) {
            fileName += ".xlsx";
        }
        if (workbook instanceof SXSSFWorkbook) {
            fileName += ".xlsx";
        }
        if (workbook instanceof HSSFWorkbook) {
            fileName += ".xls";
        }

        return fileName;
    }

    private void createHead(Sheet sheet, List<String> headList){
        CellStyle cellStyle = workbook.createCellStyle();
        createRow(sheet, headList, 0, cellStyle);
    }

    private void createBody(Sheet sheet, List<List<String>> bodyList){
        int rowSize = bodyList.size();
        CellStyle cellStyle = workbook.createCellStyle();
        for (int i = 0; i < rowSize; i++) {
            createRow(sheet, bodyList.get(i), i + 1, cellStyle);
        }
    }

    private void createRow(Sheet sheet, List<String> cellList, int rowNum, CellStyle style) {
        int size = cellList.size();
        Row row = sheet.createRow(rowNum);

        for (int i = 0; i < size; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(cellList.get(i));
        }
    }

    public static ExcelData createExcelData(List<?> data, Class<?> target, String fileName){
        return new ExcelData(
                fileName,
                createHeaderName(target),
                createBodyData(data, target)
        );
    }

    private static List<String> createHeaderName(Class<?> header){
        List<String> headData = new ArrayList<>();
        List<ExcelField> sortedColumn  = getSortedExcelColumnName(header);

        sortedColumn.forEach(column -> {
            headData.add(column.getHeaderName());
        });
        return headData;
    }

    private static List<List<String>> createBodyData(List<?> dataList, Class<?> target){
        List<List<String>> bodyData = new ArrayList<>();
        dataList.forEach(object ->{
            bodyData.add(mapToValue(object, target));
        });
        return bodyData;
    }

    private static List<String> mapToValue(Object object, Class<?> target){
        List<String> data = new ArrayList<>();
        List<ExcelField> sortedColumn  = getSortedExcelColumnName(target);
        Map<String, Object> map = objectMapper.convertValue(object, Map.class);
        sortedColumn.forEach(column -> {
                data.add(String.valueOf(map.get(column.getFieldName())));
        });
        return data;
    }

    private static List<ExcelField> getSortedExcelColumnName(Class<?> target){
        List<ExcelField> excelColumnNames = new ArrayList<>();
        for (Field field : target.getDeclaredFields()) {
            if(field.isAnnotationPresent(ExcelColumnName.class)){
                excelColumnNames.add(ExcelField.of(field));
            }
        }

        Collections.sort(excelColumnNames, new ExcelColumnNameComparator());
        return excelColumnNames;
    }
}
