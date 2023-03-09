package utils;

import cn.hutool.core.util.ObjectUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class ExcelExportUtil {
    private final static String path = "src/main/resources/";

    /**
     * feature of exporting
     *
     * @param outputStream output stream
     * @param columnList   column name list
     * @param dataList     data list
     */
    private static void export(OutputStream outputStream, List<String> columnList, List<List<String>> dataList) throws IOException {
        //excel file main object
        SXSSFWorkbook wb = null;
        try {
            //maximum number of rows in RAM, prevent out of memory error
            wb = new SXSSFWorkbook(1000);
            //create a sheet
            SXSSFSheet sheet = wb.createSheet("data");
            //row index
            Integer excelRow = 0;
            //create the first row
            SXSSFRow titleRow = sheet.createRow(excelRow++);
            if (ObjectUtil.isNotEmpty(columnList)) {
                //traversal the column name list
                for (int i = 0; i < columnList.size(); i++) {
                    // create a grid
                    Cell cell = titleRow.createCell(i);
                    // set a value for each grid
                    cell.setCellValue(columnList.get(i));
                }
            }
            if (ObjectUtil.isNotEmpty(dataList)) {
                // traversal the data list
                for (int i = 0; i < dataList.size(); i++) {
                    // create a row of data
                    Row dataRow = sheet.createRow(excelRow++);
                    // traversal every single data
                    for (int j = 0; j < dataList.get(0).size(); j++) {
                        Cell cell = dataRow.createCell(j);
                        cell.setCellValue(dataList.get(i).get(j));
                    }
                }
            }
            wb.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            try {
                if (wb != null) {
                    wb.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * reflect data export
     *
     * @param dataList
     */
    public static void reflectExport(List dataList) throws IllegalAccessException, IOException {
        //健壮性检验
        if (ObjectUtil.isEmpty(dataList)) {
            return;
        }
        List<List<String>> allList = new ArrayList<>();
        //获取class文件
        Class<?> aClass = dataList.get(0).getClass();
        //获取表名信息
        String fileName = aClass.getAnnotation(Excel.class).fileName();
        //根据文件名创建文件输出流
        OutputStream outputStream = createOutputStream(fileName);
        //通过class文件获取对象属性
        Field[] declaredFields = aClass.getDeclaredFields();
        //排除没有被注解标记的属性，函数式编程
        List<Field> fields = Arrays.asList(declaredFields).stream().filter(field -> field.isAnnotationPresent(Excel.class)).collect(Collectors.toList());
        //获取列名
        List<String> columnList = fields.stream().map(field -> {
            String columnName = field.getAnnotation(Excel.class).columnName();
            return columnName;
        }).collect(Collectors.toList());
        //属性排序
        fields = fields.stream().sorted(Comparator.comparingInt(field -> field.getAnnotation(Excel.class).order())).collect(Collectors.toList());
        //通过属性获取对象的值
        for (int i = 0; i < dataList.size(); i++) {
            List<String> list = new ArrayList<>();
            for (int j = 0; j < fields.size(); j++) {
                Field declaredField = fields.get(j);
                declaredField.setAccessible(true);
                Object obj = declaredField.get(dataList.get(i));
                if (obj != null) {
                    list.add(obj.toString());
                }
            }
            allList.add(list);
        }
        //调用第一版的导出方法
        export(outputStream, columnList, allList);
    }

    /**
     * create an output stream based on the file name. if the file doesn't exist, create a new one.
     *
     * @param fileName
     * @return
     */
    private static OutputStream createOutputStream(String fileName) {

        String filename = path + fileName + ".xlsx";
        try {
            OutputStream outputStream = new FileOutputStream(filename);
            return outputStream;
        } catch (FileNotFoundException e) {
            System.out.println("the file doesn't exist");
            return null;
        }
    }
}
