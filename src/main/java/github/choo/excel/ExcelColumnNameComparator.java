package github.choo.excel;

import java.util.Comparator;

public class ExcelColumnNameComparator implements Comparator<ExcelField> {
    @Override
    public int compare(ExcelField o1, ExcelField o2) {
        if(o1.getOrder() > o2.getOrder()){
            return 1;
        }
        else if(o1.getOrder() < o2.getOrder()){
            return -1;
        }
        return 0;
    }
}
