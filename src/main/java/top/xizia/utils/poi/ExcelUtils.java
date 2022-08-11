package top.xizia.utils.poi;


import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.net.URLEncoder;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @NAME: WSC
 * @DATE: 2021/12/2
 * @DESCRIBE:
 **/
public class ExcelUtils {

    private static int maxHeadRow = 1;
    private static int startIndex;
    private static Map<String, List<Field>> parentChildFieldMap = null;
    private static Map<String, Excel> excelMap = null;

    private static Map<String, Object> emptyInstanceCache = null;

    public static <T> void realDownloadExcel(List<T> arr, HttpServletResponse response) throws IllegalAccessException, IOException, InstantiationException {
        maxHeadRow = 1;
        startIndex = -1;

        if (arr != null && arr.size() != 0) {
            List<List<String>> dataList = new ArrayList<>();

            Field[] fields = arr.get(0).getClass().getDeclaredFields();
            Field[] parentFields = arr.get(0).getClass().getSuperclass().getDeclaredFields();

            List<Field> mainFields = new ArrayList<>();
            List<String> columnList = new ArrayList<>();
            excelMap = new HashMap<>();
            List<List<Object>> mergeColumnInfo = new ArrayList<>();
            parentChildFieldMap = new HashMap<>();
            emptyInstanceCache = new HashMap<>();

            /**
             * 获取父类的所有字段
             */
            for (int i = 0; i < parentFields.length; i++) {
                Field field = parentFields[i];
                if (field.isAnnotationPresent(Excel.class)) {
                    startIndex++;
                    field.setAccessible(true);
                    mainFields.add(field);
                    excelMap.put(field.getName(), field.getAnnotation(Excel.class));
                    int cnt = resignMultiHeadFieldInfoIfNecessary(field, 1, mainFields, excelMap, mergeColumnInfo, startIndex);
                    startIndex += (cnt == 0) ? 0 : cnt - 1;
                }
            }

            /**
             * 获取子类的字段
             */
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                if (field.isAnnotationPresent(Excel.class)) {
                    startIndex++;
                    field.setAccessible(true);
                    mainFields.add(field);
                    excelMap.put(field.getName(), field.getAnnotation(Excel.class));
                    int cnt = resignMultiHeadFieldInfoIfNecessary(field, 1, mainFields, excelMap, mergeColumnInfo, startIndex);
                    startIndex += (cnt == 0) ? 0 : cnt - 1;
                }
            }


            /**
             * 排序Sort
             */
            List<Field> fieldList = mainFields.stream()
                    .sorted((o1, o2) -> {
                        boolean b1 = o1.isAnnotationPresent(Excel.class);
                        boolean b2 = o2.isAnnotationPresent(Excel.class);

                        if (b1 && b2) {
                            Excel o1Annotation = o1.getAnnotation(Excel.class);
                            Excel o2Annotation = o2.getAnnotation(Excel.class);

                            return o1Annotation.sort() - o2Annotation.sort();
                        }

                        return b1 ? -1 : 1;
                    }).collect(Collectors.toList());


            mergeNotMultiFiledColumn(1, 0, mergeColumnInfo, mainFields);

            /**
             * 获取列的字段名称
             */
            for (int i = 0; i < fieldList.size(); i++) {
                Field field = fieldList.get(i);
                Excel excel = field.getAnnotation(Excel.class);
                if (excel.isMultipleHeaders()) {
                    List<Field> list = parentChildFieldMap.get(field.getName());
                    for (Field field1 : list) {
                        Excel excel1 = excelMap.get(field1.getName());
                        if (excel1.isMultipleHeaders()) {
                            readMultiFiledColumn(field1, columnList);
                        } else {
                            columnList.add(excel1.value());
                        }
                    }
                } else {
                    columnList.add(excel.value());
                }
            }

            int idx = 0;

            for (T t : arr) {
                List<String> rowData = new ArrayList<>();
                idx++;

                for (int i = 0; i < fieldList.size(); i++) {
                    Field field = fieldList.get(i);

                    Excel excel = excelMap.get(field.getName());

                    if (excel != null && excel.isIndex()) {
                        /**
                         * 表头是序号
                         */
                        rowData.add(idx + "");
                    } else if (excel.isMultipleHeaders()) {
                        /**
                         * 多级表头
                         */
                        readMultiFieldInfo(field, rowData, t);

                    } else {
                        Object fValue = field.get(t);
                        Object value = isObjectEmpty(fValue) ? "" : fValue;
                        rowData.add(value + "");
                    }
                }

                dataList.add(rowData);
            }

            /**
             * 淘汰掉 性能太低了！
             */
//            Workbook wb = POIUtils.createWorkBook("2007", null, columnList, dataList);

            ServletOutputStream outputStream = response.getOutputStream();
            ExcelWriter writer = ExcelUtil.getWriter(true);

            /**
             * 合并表头
             */
            for (List<Object> list : mergeColumnInfo) {
                writer.merge((Integer) list.get(0)
                        , (Integer) list.get(1)
                        , (Integer) list.get(2)
                        , (Integer) list.get(3)
                        , list.get(4), true);
            }

            writer.passRows(maxHeadRow - 1);
            writer.writeHeadRow(columnList);
            writer.write(dataList);
            writer.flush(outputStream, true);
            writer.close();
        }
    }

    /**
     * 合并非多表头的表格
     *
     * @param currentRow
     * @param mergeColumnInfo
     * @param mainFields
     */
    public static int mergeNotMultiFiledColumn(int currentRow, int startIndex, List<List<Object>> mergeColumnInfo, List<Field> mainFields) {
        int currIndex = startIndex;

        for (int i = 0; i < mainFields.size(); i++) {
            Field field = mainFields.get(i);
            Excel excel = excelMap.get(field.getName());

            if (!excel.isMultipleHeaders()) {
                if (maxHeadRow - currentRow > 0) {
                    mergeColumnInfo.add(Arrays.asList(currentRow - 1, maxHeadRow - 1, currIndex, currIndex, excel.value()));
                }
            } else {
                List<Field> fields = parentChildFieldMap.get(field.getName());
                currIndex = mergeNotMultiFiledColumn(currentRow + 1, currIndex, mergeColumnInfo, fields);
            }

            currIndex++;
        }

        return currIndex - 1;
    }

    public static void readMultiFiledColumn(Field parentFiled, List<String> columnList) {
        List<Field> fields = parentChildFieldMap.get(parentFiled.getName());

        for (Field field : fields) {
            Excel excel = excelMap.get(field.getName());
            if (excel.isMultipleHeaders()) {
                readMultiFiledColumn(field, columnList);
            } else {
                columnList.add(excel.value());
            }
        }
    }

    public static void readMultiFieldInfo(Field parentFiled, List<String> rowData, Object instant) throws IllegalAccessException, InstantiationException {

        /**
         * 多级表头
         */
        List<Field> list = parentChildFieldMap.get(parentFiled.getName());
        Object parentObject = null;
        if (instant != null) {
            parentObject = parentFiled.get(instant);
        }


        ListIterator<Field> iterator = list.listIterator();
        while (iterator.hasNext()) {
            Field field1 = iterator.next();

            Excel excel = excelMap.get(field1.getName());

            if (excel.isMultipleHeaders()) {
                if (parentObject != null) {
                    readMultiFieldInfo(field1, rowData, parentObject);
                } else {
                    Object instance = emptyInstanceCache.get(field1.getDeclaringClass().getName());

                    if (instance == null) {
                        field1.setAccessible(true);
                        instance = field1.getDeclaringClass().newInstance();
                        emptyInstanceCache.put(field1.getDeclaringClass().getName(), instance);
                    }

                    readMultiFieldInfo(field1, rowData, instance);
                }
            } else {
                if (isObjectEmpty(parentObject)) {
                    rowData.add("");
                } else {
                    Object fValue = field1.get(parentObject);
                    Object value = isObjectEmpty(fValue) ? "" : fValue;
                    rowData.add(value + "");
                }
            }
        }
    }

    /**
     * 注册多表头信息
     *
     * @param parentHead 父字段
     * @param mainFields 所有的Field集合
     * @param excelMap   所有的相关注解的集合
     */
    public static int resignMultiHeadFieldInfoIfNecessary(Field parentHead
            , int parentCurrentRow
            , List<Field> mainFields
            , Map<String, Excel> excelMap
            , List<List<Object>> mergeColumnInfo
            , int mStartIndex) {

        Excel excel = parentHead.getAnnotation(Excel.class);

        if (!excel.isMultipleHeaders()) {
            return 0;
        }

        int cnt = 0;

        List<Field> childFields = new ArrayList<>();

        Field[] fields = parentHead.getType().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            Excel childExcel = field.getAnnotation(Excel.class);
            if (childExcel != null) {
                childFields.add(field);
                excelMap.put(field.getName(), childExcel);

                if (childExcel.isMultipleHeaders()) {
                    cnt += resignMultiHeadFieldInfoIfNecessary(field, parentCurrentRow + 1, mainFields, excelMap, mergeColumnInfo, mStartIndex + cnt);
                } else {
                    cnt++;
                }
            }
        }

        int endIndex = mStartIndex + cnt - 1;

        /**
         * 不能等于, 等于的话只有一个列头的话 那多列头已经没有意义
         */
        if (startIndex < endIndex) {
            if (maxHeadRow < parentCurrentRow + 1) {
                maxHeadRow = parentCurrentRow + 1;
            }

            parentChildFieldMap.put(parentHead.getName(), childFields);

            // 记录合并信息
            mergeColumnInfo.add(Arrays.asList(parentCurrentRow - 1, parentCurrentRow - 1, mStartIndex, endIndex, excel.value()));
        }

        return cnt;
    }


    public static <T> List<T> readExcel(Class<T> clazz, String fileName) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        List<T> arr = new ArrayList<>();

        ExcelReader reader = ExcelUtil.getReader(fileName);
        List<Map<String, Object>> list = reader.readAll();

        if (list != null && list.size() != 0) {

            Map<String, String> fieldMatch = new HashMap<>();

            T t = clazz.getDeclaredConstructor().newInstance();
            Field[] fields = t.getClass().getDeclaredFields();
            Field[] parentFields = t.getClass().getSuperclass().getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(Excel.class)) {
                    Excel excel = field.getAnnotation(Excel.class);

                    if (!ObjectUtil.isEmpty(excel.value())) {
                        fieldMatch.put(excel.value(), field.getName());
                    }
                }
            }

            for (Field field : parentFields) {
                if (field.isAnnotationPresent(Excel.class)) {
                    Excel excel = field.getAnnotation(Excel.class);

                    if (!ObjectUtil.isEmpty(excel.value())) {
                        fieldMatch.put(excel.value(), field.getName());
                    }
                }
            }


            for (Map<String, Object> map : list) {
                Map<String, String> properties = new HashMap<>();

                for (Map.Entry<String, Object> entry : map.entrySet()) {
                    String key = fieldMatch.getOrDefault(entry.getKey(), entry.getKey());
                    String value = entry.getValue().toString();

                    properties.put(key, value);
                }

                T item = BeanUtil.mapToBean(properties, clazz, true);
                arr.add(item);
            }


        }

        return arr;
    }

    public static <T> void downloadExcel(HttpServletResponse response, List<T> list) throws IOException, IllegalAccessException, InstantiationException {
        long startTime = System.currentTimeMillis();

        String fileName = System.currentTimeMillis() + ".xlsx";
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));

        ExcelUtils.realDownloadExcel(list, response);


        long spendTime = System.currentTimeMillis() - startTime;
        System.out.println("导出和下载报表,一共花费了:" + spendTime + "ms");
    }

    private static boolean isObjectEmpty(Object o) {
        return ObjectUtil.isEmpty(o) ||
                (o instanceof CharSequence && o.toString().equalsIgnoreCase("null"));
    }
}
