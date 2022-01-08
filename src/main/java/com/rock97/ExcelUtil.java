package com.rock97;

import com.rock97.annotation.Excel;
import com.rock97.annotation.ExcelEnum;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @Description: 导出工具类
 * @Author: lizhihua
 * @Email: lizhihua@mobike.com
 * @Create: 2018-04-18 15:54
 */
public class ExcelUtil {

    //标题名
    private String title;
    //sheet名数组
    private String sheetNames;
    //列名数组
    private String[] columnNames;
    //列对应实体名数组
    private String[] entityNames;
    //数据
    private int sum = 0;
    private SXSSFWorkbook workbook;
    private Sheet sheet;
    private String classNames;
    private static final ConcurrentHashMap<String, Method> methodCache = new ConcurrentHashMap();
    private static final ConcurrentHashMap<String, Field> fieldCache = new ConcurrentHashMap();
    private static final LinkedHashMap<String, String> excelEunmCache = new LinkedHashMap();
    private ExcelUtil(LinkedHashMap<String, String> map, String title, String classNames){
        workbook = new SXSSFWorkbook(200);
        this.classNames = classNames;
        this.sheetNames = title;
        this.title = title;
        this.sheet = workbook.createSheet(sheetNames);
        int column = 0;

        columnNames = new String[map.size()];
        entityNames = new String[map.size()];
        for(Iterator var6 = map.keySet().iterator(); var6.hasNext(); ++column) {
            String key = (String)var6.next();
            columnNames[column] = map.get(key);
            entityNames[column] = key;
        }
        //设置标题
        setSheetitle();
    }

    /**
     * 获取导出工具类实例对象
     * @param type
     * @param title
     * @return
     */
    public static ExcelUtil newInstance(Class type, String title){
        String className = type.getName();
        Field[] fields = type.getDeclaredFields();
        LinkedHashMap<String, String> map = new LinkedHashMap();
        LinkedHashMap<String, String> mapExcel = new LinkedHashMap();
        Field[] var5 = fields;
        int var6 = fields.length;
        List<Excel> excelList = new ArrayList<>();

        for(int var7 = 0; var7 < var6; ++var7) {
            Field field = var5[var7];
            if (field.isAnnotationPresent(Excel.class)) {
                Excel excel = field.getAnnotation(Excel.class);
                mapExcel.put(excel.value(),field.getName());
                excelList.add(excel);
            }
            if(field.isAnnotationPresent(ExcelEnum.class)){
                ExcelEnum excelEnum = field.getAnnotation(ExcelEnum.class);
                String[] codes = excelEnum.code();
                String[] names = excelEnum.name();
                if (codes != null) {
                    for (int i = 0; i < codes.length; i++) {
                        String key = className + field.getName()+codes[i];
                        excelEunmCache.put(key,names[i]);
                    }
                }
            }
        }
        excelList.sort((h1, h2) -> h1.index() - h2.index());
        for (Excel excel : excelList) {
            map.put(mapExcel.get(excel.value()),excel.value());
        }
        return new ExcelUtil(map,title,className);
    }

    /**
     * 分批添加元素
     * @param list
     */
    public void addList(List list){
        List<Object[]> datas = new ArrayList<Object[]>();
        if(CollectionUtils.isEmpty(list)){
            list = new ArrayList();
        }
        for (Object t : list) {
            datas.add(beanToArray(t, entityNames));
        }
        createWorkBook(datas,sum);
        sum += datas.size();
    }

    /**
     * 获取导出数据
     * @return
     */
    private byte[] exportExcel(){
        byte[] bytes = null;
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            workbook.write(out);
            //清理磁盘文件
            workbook.dispose();
            bytes = out.toByteArray();

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        return bytes;
    }

    /**
     * 写入流
     * @param out
     */
    public void write(OutputStream out){
        try {
            workbook.write(out);
            //清理磁盘文件
            workbook.dispose();

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
    }


    private void setSheetitle() {
        int colunmSize = columnNames.length;
        //设置列名
        Row columnRow = sheet.createRow(0);
        for (int i = 0; i < colunmSize; i++) {
            Cell columnRowCell = columnRow.createCell(i, CellType.STRING);
            columnRowCell.setCellValue(columnNames[i]);
        }
    }
    /**
     * 组装excel
     * @param datas
     * @param begin
     */
    private void createWorkBook(List<Object[]> datas,int begin) {

        //设置数据
        for (int i = 0; i < datas.size(); i++) {
            Object[] obj = datas.get(i);
            Row dataRow = sheet.createRow(i + 1 + begin);
            for (int j = 0; j < obj.length; j++) {
                Cell dataRowCell = dataRow.createCell(j, CellType.STRING);
                if (null != obj[j] && obj[j].toString()!=null && !"".equals(obj[j].toString())) {
                    dataRowCell.setCellValue(obj[j].toString());
                } else {
                    dataRowCell.setCellValue("");
                }
              //  dataRowCell.setCellStyle(style);
            }
        }

    }

    /**
     * javaBean转为Object[]
     */
    private Object[] beanToArray(Object t, String[] entityNames) {
        SimpleDateFormat SIMPLE_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Object[] objArray = new Object[entityNames.length];
        try {
            for (int i = 0; i < entityNames.length; i++) {
                String name = "get"+entityNames[i].substring(0, 1).toUpperCase() + entityNames[i].substring(1);
                String key = classNames+entityNames[i];
                Method m = methodCache.get(key);
                if(null == m){
                    m = t.getClass().getMethod(name);
                    methodCache.put(key,m);
                }

                Field field = fieldCache.get(key);
                if(field == null){
                    field = t.getClass().getDeclaredField(entityNames[i]);
                    fieldCache.put(key,field);
                }

                Object object = m.invoke(t);

                if(field.isAnnotationPresent(ExcelEnum.class)){
                    objArray[i] = excelEunmCache.get(key + object);
                }else if(object!=null && object instanceof Date){
                    objArray[i] = SIMPLE_DATE_FORMAT.format(object);
                }else{
                    objArray[i] =object;
                }
            }
        } catch (Exception e) {
           e.printStackTrace();
        }
        return objArray;
    }
}