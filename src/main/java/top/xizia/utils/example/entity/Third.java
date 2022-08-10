package top.xizia.utils.example.entity;

import top.xizia.utils.poi.Excel;

/**
 * @author: WSC
 * @DATE: 2022/8/10
 * @DESCRIBE:
 **/
public class Third {
    @Excel("测试1")
    private String test1;

    @Excel("测试2")
    private String test2;

    @Excel(value = "第四", isMultipleHeaders = true)
    private Fourth fourth;

    public String getTest1() {
        return test1;
    }

    public void setTest1(String test1) {
        this.test1 = test1;
    }

    public String getTest2() {
        return test2;
    }

    public void setTest2(String test2) {
        this.test2 = test2;
    }

    public Fourth getFourth() {
        return fourth;
    }

    public void setFourth(Fourth fourth) {
        this.fourth = fourth;
    }
}
