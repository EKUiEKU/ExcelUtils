package top.xizia.utils.poi;

import cn.hutool.poi.excel.ExcelWriter;

import java.util.List;

/**
 * @author: WSC
 * @DATE: 2022/8/19
 * @DESCRIBE: ExcelUtil模板样式监听器
 **/
public interface TemplateStyleListener {

    /**
     * 往表头添加信息
     */
    List<List<List<Object>>> onAddHeadInfoEvent();


    /**
     * 往表底添加信息
     * @param left
     * @param right
     * @param top
     * @param bottom
     */
    List<List<List<Object>>> onAddBottomInfoEvent(int left, int right, int top, int bottom);


    /**
     * 导出报表的时候 可以设置样式的补充一些信息
     * @param writer
     * @param left
     * @param right
     * @param top
     * @param bottom
     */
    void onModifyStyleEvent(ExcelWriter writer, int left, int right, int top, int bottom);
}
