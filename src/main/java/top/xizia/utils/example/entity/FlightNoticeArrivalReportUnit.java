package top.xizia.utils.example.entity;

import top.xizia.utils.poi.Excel;

/**
 * @author: WSC
 * @DATE: 2022/8/8
 * @DESCRIBE: 进港播报查询报表详情的多级表头DTO
 **/
public class FlightNoticeArrivalReportUnit {
    /**
     * 用时
     */
    @Excel(value = "用时", width = 10)
    private Long spendTime;
    /**
     * 时限
     */
    @Excel(value = "时限", sort = 1, width = 10)
    private Long dateLine;
    /**
     * 超时时间
     */
    @Excel(value = "超时时间",sort = 2, width = 10)
    private Long overtimeTime;

    @Excel(value = "第三", isMultipleHeaders = true, width = 10)
    private Third third;


    public Long getSpendTime() {
        return spendTime;
    }

    public void setSpendTime(Long spendTime) {
        this.spendTime = spendTime;
    }

    public Long getDateLine() {
        return dateLine;
    }

    public void setDateLine(Long dateLine) {
        this.dateLine = dateLine;
    }

    public Long getOvertimeTime() {
        return overtimeTime;
    }

    public void setOvertimeTime(Long overtimeTime) {
        this.overtimeTime = overtimeTime;
    }

    public Third getThird() {
        return third;
    }

    public void setThird(Third third) {
        this.third = third;
    }
}
