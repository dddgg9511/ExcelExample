package github.choo.excel;


import github.choo.excel.annotation.ExcelColumnName;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

import java.lang.reflect.Field;

@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public class ExcelField {
    private int order;
    private String headerName;
    private String fieldName;

    public static ExcelField of(Field field){
        ExcelColumnName annotation = field.getAnnotation(ExcelColumnName.class);
        return new ExcelField(annotation.order(), annotation.headerName(), field.getName());
    }
}
