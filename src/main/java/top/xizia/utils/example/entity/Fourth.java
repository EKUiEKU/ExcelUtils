package top.xizia.utils.example.entity;

import top.xizia.utils.poi.Excel;

/**
 * @author: WSC
 * @DATE: 2022/8/10
 * @DESCRIBE:
 **/
public class Fourth {
    @Excel("测试三")
    private String test3;

    @Excel("测试四")
    private String test4;

    @Excel("测试五")
    private String test5;

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
