package top.xizia.utils.poi;


import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.math.RoundingMode;
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

    private static int startRowIndex;

    private static List<Integer> columnWidthList = null;

    private static Map<Long, BigDecimal> calculateMap = null;
    private static Map<String, Aggregation> aggregationMap = null;
    private static Map<Long, Aggregation> aggregationIndexMap = null;

    private static String excelTitle = "导出报表大标题";

    public static <T> void realDownloadExcel(List<T> arr, HttpServletResponse response, String title, TemplateStyleListener templateStyleListener) throws IllegalAccessException, IOException, InstantiationException {
        startRowIndex = 2;

        List<List<List<Object>>> addHeadInfo = null;

        if (templateStyleListener != null) {
            addHeadInfo = templateStyleListener.onAddHeadInfoEvent();
            if (addHeadInfo != null) {
                startRowIndex += addHeadInfo.size();
            }
        }

        maxHeadRow = startRowIndex;
        startIndex = -1;

        if (title == null) {
            maxHeadRow --;
            startRowIndex --;
        }

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
            columnWidthList = new ArrayList<>(16);
            aggregationMap = new HashMap<>();
            calculateMap = new HashMap<>();
            aggregationIndexMap = new HashMap<>();

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
                    resignAggregationIfNecessary(field);
                    int cnt = resignMultiHeadFieldInfoIfNecessary(field, startRowIndex + 1, mainFields, excelMap, mergeColumnInfo, startIndex);
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
                    resignAggregationIfNecessary(field);
                    int cnt = resignMultiHeadFieldInfoIfNecessary(field, startRowIndex + 1, mainFields, excelMap, mergeColumnInfo, startIndex);
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


            mergeNotMultiFiledColumn(startRowIndex + 1, 0, mergeColumnInfo, mainFields);

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
                            columnWidthList.add(excel1.width());
                        }
                    }
                } else {
                    columnList.add(excel.value());
                    columnWidthList.add(excel.width());
                }
            }

            int idx = 0;

            for (T t : arr) {
                List<String> rowData = new ArrayList<>();
                idx++;

                for (int i = 0; i < fieldList.size(); i++) {
                    Field field = fieldList.get(i);

                    Excel excel = excelMap.get(field.getName());

                    if (excel != null) {
                        Aggregation aggregation = aggregationMap.get(field.getName());

                        if (excel.isIndex()) {
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

                            if (fValue == null) {
                                fValue = emptyInstanceCache.get(field.getName());
                            }

                            if (aggregation != null && fValue instanceof Number) {
                                if (!aggregationIndexMap.containsKey(Long.valueOf(i))) {
                                    aggregationIndexMap.put(Long.valueOf(i), aggregation);
                                }

                                BigDecimal decimal = calculateMap.get(Long.valueOf(i));

                                BigDecimal tmpNumber = null;
                                if (fValue instanceof BigDecimal) {
                                    tmpNumber = (BigDecimal) fValue;
                                } else {
                                    tmpNumber = BigDecimal.valueOf(Long.valueOf(fValue + ""));
                                }

                                if (decimal == null) {
                                    calculateMap.put(Long.valueOf(i), tmpNumber);
                                } else {
                                    decimal = decimal.add(tmpNumber);
                                    calculateMap.put(Long.valueOf(i), decimal);
                                }
                            }

                            Object value = isObjectEmpty(fValue) ? "" : fValue;
                            rowData.add(value + "");
                        }
                    }
                }

                dataList.add(rowData);
            }



            String[] sumCalculate = new String[columnList.size()];
            String[] avgCalculate = new String[columnList.size()];
            Boolean hasSumCalculate = false;
            Boolean hasAvgCalculate = false;
            if (calculateMap.size() != 0) {
                sumCalculate[0] = "总计";
                Set<Map.Entry<Long, BigDecimal>> entries = calculateMap.entrySet();
                for (Map.Entry<Long, BigDecimal> entry : entries) {
                    int i = entry.getKey().intValue();
                    Aggregation aggregation = aggregationIndexMap.get(Long.valueOf(i));
                    if (aggregation.equals(Aggregation.SUM)) {
                        sumCalculate[entry.getKey().intValue()] = entry.getValue().toString();
                        hasSumCalculate = true;
                    }else if (aggregation.equals(Aggregation.AVG)) {
                        avgCalculate[entry.getKey().intValue()] = entry.getValue().toString();
                        hasAvgCalculate = true;
                    }else if(aggregation.equals(Aggregation.BOTH)) {
                        sumCalculate[entry.getKey().intValue()] = entry.getValue().toString();
                        avgCalculate[entry.getKey().intValue()] = entry.getValue().divide(BigDecimal.valueOf(dataList.size()), RoundingMode.HALF_UP).toString();
                        hasSumCalculate = true;
                        hasAvgCalculate = true;
                    }
                }

                if (hasSumCalculate) {
                    sumCalculate[0] = "总计";
                    dataList.add(Arrays.asList(sumCalculate));
                }

                if (hasAvgCalculate) {
                    avgCalculate[0] = "平均值";
                    dataList.add(Arrays.asList(avgCalculate));
                }
            }


            /**
             * 淘汰掉 性能太低了！
             */
//            Workbook wb = POIUtils.createWorkBook("2007", null, columnList, dataList);

            ServletOutputStream outputStream = response.getOutputStream();
            ExcelWriter writer = ExcelUtil.getWriter(true);

            /**
             * 大标题
             */
            XSSFFont bigTitleFont = new XSSFFont();
            bigTitleFont.setFontName("黑体");
            bigTitleFont.setFontHeightInPoints((short) 18);

            XSSFRichTextString bigTitleString = new XSSFRichTextString(excelTitle);
            bigTitleString.applyFont(bigTitleFont);

            writer.merge(0, 0, 0, columnList.size() - 1, bigTitleString, false);

            /**
             * 把线隐藏掉
             */
            CellStyle cellStyle = writer.createCellStyle();
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setRightBorderColor((short) 0xD4D4D4);

            for (int i = 0; i < columnList.size(); i++) {
                writer.setStyle(cellStyle, i, 0);
            }


            /**
             * 小标题
             */
            if (title != null) {
                XSSFFont secondTitleFont = new XSSFFont();
                secondTitleFont.setFontName("黑体");
                secondTitleFont.setFontHeightInPoints((short) 14);

                XSSFRichTextString secondTitleString = new XSSFRichTextString(title);
                secondTitleString.applyFont(secondTitleFont);


                writer.merge(1, 1, 0, columnList.size() - 1, secondTitleString, false);

                for (int i = 0; i < columnList.size(); i++) {
                    writer.setStyle(cellStyle, i, 1);
                }
            }

            /**
             * 合并补充的信息
             */
            int tempCurrentRowIndex = 2;
            if (addHeadInfo != null) {
                for (List<List<Object>> row : addHeadInfo) {
                    for (List<Object> objects : row) {
                        // startX width text
                        int startX  = (int) objects.get(0);
                        int width  = (int) objects.get(1);
                        Object content = objects.get(2);

                        if (width == 1) {
                            writer.writeCellValue(startX, tempCurrentRowIndex, content);
                            writer.setStyle(null, startX, tempCurrentRowIndex);
                        }else {
                            writer.merge(tempCurrentRowIndex
                                    , tempCurrentRowIndex
                                    , startX
                                    , startX + width - 1
                                    , content, false);
                            writer.setStyle(null, startX, tempCurrentRowIndex);

                            for (int i = startX; i <= startX + width; i++) {
                                writer.setStyle(null, i, tempCurrentRowIndex);
                            }
                        }
                    }

                    tempCurrentRowIndex++;
                }
            }

            /**
             * 合并表头
             */
            for (List<Object> list : mergeColumnInfo) {
                writer.merge((Integer) list.get(0)
                        , (Integer) list.get(1)
                        , (Integer) list.get(2)
                        , (Integer) list.get(3)
                        , list.get(4), false);
            }

            writer.passRows(maxHeadRow - 1);
            writer.writeHeadRow(columnList);

            writer.write(dataList);

            /**
             * 设置表头宽度
             */
            for (int i = 0; i < columnWidthList.size(); i++) {
                Integer width = columnWidthList.get(i);

                if (width <= 0) {
                    writer.autoSizeColumn(i);
                }else {
                    if (width > 255) {
                        width = 255;
                    }

                    writer.setColumnWidth(i, width);
                }
            }

            // writer.setFreezePane(maxHeadRow + dataList.size() + 1, columnList.size());



            // // 制表人
            // writer.writeCellValue(0, maxHeadRow + dataList.size() + 2, "制表人");
            // writer.setStyle(null, 0, maxHeadRow + dataList.size() + 2);
            // writer.writeCellValue(1, maxHeadRow + dataList.size() + 2, UserInfoUtil.getUserAccount());
            // writer.setStyle(null, 1, maxHeadRow + dataList.size() + 2);
            // // 角色
            // String roleNames = UserInfoUtil.getUserRoleNameList()
            //         .stream()
            //         .collect(Collectors.joining(","));
            // writer.writeCellValue(0, maxHeadRow + dataList.size() + 3, roleNames);
            // writer.setStyle(null, 0, maxHeadRow + dataList.size() + 3);
            tempCurrentRowIndex = maxHeadRow + dataList.size();
            if(templateStyleListener != null) {
                List<List<List<Object>>> addBottomInfo = templateStyleListener.onAddBottomInfoEvent(0, columnList.size() - 1, 0, maxHeadRow + dataList.size());
                if (addBottomInfo != null) {
                    for (List<List<Object>> row : addBottomInfo) {
                        for (List<Object> objects : row) {
                            // startX width text
                            int startX  = (int) objects.get(0);
                            int width  = (int) objects.get(1);
                            Object content = objects.get(2);

                            if (width == 1) {
                                writer.writeCellValue(startX, tempCurrentRowIndex, content);
                                writer.setStyle(null, startX, tempCurrentRowIndex);
                            }else {
                                writer.merge(tempCurrentRowIndex
                                        , tempCurrentRowIndex
                                        , startX
                                        , startX + width - 1
                                        , content, false);

                                for (int i = startX; i <= startX + width; i++) {
                                    writer.setStyle(null, i, tempCurrentRowIndex);
                                }
                            }
                        }

                        tempCurrentRowIndex++;
                    }
                }
            }

            /**
             * 触发监听器
             */
            if (templateStyleListener != null) {
                templateStyleListener.onModifyStyleEvent(writer, 0, columnList.size() - 1, 0, maxHeadRow + dataList.size());
            }

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
                columnWidthList.add(excel.width());
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
            , int currentCurrentRow
            , List<Field> mainFields
            , Map<String, Excel> excelMap
            , List<List<Object>> mergeColumnInfo
            , int mStartIndex) {

        Excel excel = parentHead.getAnnotation(Excel.class);

        if (!excel.isMultipleHeaders()) {
            maxHeadRow = Math.max(maxHeadRow, currentCurrentRow);
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
                resignAggregationIfNecessary(field);
                if (childExcel.isMultipleHeaders()) {
                    cnt += resignMultiHeadFieldInfoIfNecessary(field, currentCurrentRow + 1, mainFields, excelMap, mergeColumnInfo, mStartIndex + cnt);
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
            maxHeadRow = Math.max(maxHeadRow, currentCurrentRow + 1);

            parentChildFieldMap.put(parentHead.getName(), childFields);

            // 记录合并信息
            mergeColumnInfo.add(Arrays.asList(currentCurrentRow - 1, currentCurrentRow - 1, mStartIndex, endIndex, excel.value()));
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

    public static <T> void downloadExcel(HttpServletResponse response, List<T> list, String secondTitle, String bigTitle, TemplateStyleListener templateStyleListener) throws IOException, IllegalAccessException, InstantiationException {
        long startTime = System.currentTimeMillis();

        String fileName = System.currentTimeMillis() + ".xlsx";
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));

        if (list == null || list.size() == 0) {
            throw new IllegalAccessException("报表数据不能为空！");
        }

        if (bigTitle != null) {
            excelTitle = bigTitle;
        }

        ExcelUtils.realDownloadExcel(list, response, secondTitle, templateStyleListener);


        long spendTime = System.currentTimeMillis() - startTime;
        System.out.println("导出和下载报表,一共花费了:" + spendTime + "ms");
    }

    public static <T> void downloadExcel(HttpServletResponse response, List<T> list, String secondTitle, String bigTitle) throws IOException, IllegalAccessException, InstantiationException {
        downloadExcel(response, list, secondTitle, bigTitle, null);
    }

    public static <T> void downloadExcel(HttpServletResponse response, List<T> list, String secondTitle) throws IOException, IllegalAccessException, InstantiationException {
        downloadExcel(response, list, secondTitle, null, null);
    }

    public static <T> void downloadExcel(HttpServletResponse response, List<T> list) throws IOException, IllegalAccessException, InstantiationException {
        downloadExcel(response, list, null, null, null);
    }

    private static boolean isObjectEmpty(Object o) {
        return ObjectUtil.isEmpty(o) ||
                (o instanceof CharSequence && o.toString().equalsIgnoreCase("null"));
    }

    public static void resignAggregationIfNecessary(Field field) {
        Excel excel = excelMap.get(field.getName());
        if (excel != null && !excel.aggregation().equals(Aggregation.NONE)) {
            aggregationMap.put(field.getName(), excel.aggregation());
        }
    }
}
