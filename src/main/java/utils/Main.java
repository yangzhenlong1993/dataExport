package utils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException, IllegalAccessException {
        Student student1 = new Student("ivy", 18);
        Student student2 = new Student("claude", 30);
        List<Student> list = new ArrayList<>();
        list.add(student1);
        list.add(student2);
        ExcelExportUtil.reflectExport(list);
    }
}
