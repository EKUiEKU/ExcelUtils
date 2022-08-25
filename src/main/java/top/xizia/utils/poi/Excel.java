package top.xizia.utils.poi;

import java.lang.annotation.*;

/**
 * @author wsc
 * @NAME: WSC
 * @DATE: 2021/12/2
 * @DESCRIBE:
 **/
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface Excel {
    String value() default "";

    /**
     * 排序
     * @return
     */
    int sort() default 0;

    /**
     * 是否为序号
     * @return
     */
    boolean isIndex() default false;

    /**
     * 是否为多表格, 要求字段是一个Object
     * @return
     */
    boolean isMultipleHeaders() default false;

    /**
     * 表格宽度
     * @return
     */
    int width() default 10;

    /**
     * 总计/平均值
     * @return
     */
    Aggregation aggregation() default Aggregation.NONE;
}
