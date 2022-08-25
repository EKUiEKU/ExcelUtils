package top.xizia.utils.example.entity;

import top.xizia.utils.poi.Excel;

/**
 * @author: WSC
 * @DATE: 2022/8/10
 * @DESCRIBE:
 **/
public class Fourth {
    @Excel(value = "测试三", width = 10)
    private String test3;

    @Excel(value = "测试四", width = 10)
    private String test4;

    @Excel(value = "测试五", width = 10)
    private String test5;


    @Excel(value = "第五", isMultipleHeaders = true, width = 10)
    private Five five;

    @Excel(value = "测试六", width = 10)
    private String test6;


    public String getTest6() {
        return test6;
    }


    public Five getFive() {
        return five;
    }

    public void setFive(Five five) {
        this.five = five;
    }

    public void setTest6(String test6) {
        this.test6 = test6;
    }

    public String getTest3() {
        return test3;
    }

    public void setTest3(String test3) {
        this.test3 = test3;
    }

    public String getTest4() {
        return test4;
    }

    public void setTest4(String test4) {
        this.test4 = test4;
    }

    public String getTest5() {
        return test5;
    }

    public void setTest5(String test5) {
        this.test5 = test5;
    }
}
