package top.xizia.utils.example.entity;

import top.xizia.utils.poi.Excel;

/**
 * @author: WSC
 * @DATE: 2022/8/11
 * @DESCRIBE:
 **/
public class Five {
    @Excel(value = "你好", width = 10)
    private String hello;

    @Excel(value = "世界", width = 10)
    private String world;
}
