package github.choo.excel.data;

import java.util.List;

public class ExcelData {
    private String fileName;
    private List<String> headers;
    private List<List<String>> bodies;

    public String getFileName() {
        return fileName;
    }

    public List<String> getHeaders() {
        return headers;
    }

    public List<List<String>> getBodies() {
        return bodies;
    }

    public ExcelData(String fileName, List<String> headers, List<List<String>> bodies) {
        this.fileName = fileName;
        this.headers = headers;
        this.bodies = bodies;
    }
}
