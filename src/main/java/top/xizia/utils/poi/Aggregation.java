package top.xizia.utils.poi;

/**
 * @author: WSC
 * @DATE: 2022/8/11
 * @DESCRIBE:
 **/
public enum Aggregation {
    NONE(0, ""),
    /**
     * 总计
     */
    SUM(1, "总计"),
    /**
     * 平均值
     */
    AVG(2, "平均值"),
    /**
     * 总计和平均值
     */
    BOTH(3, "总计和平均值");
    /**
     * 聚合编码
     */
    private Integer code;
    /**
     * 聚合名称
     */
    private String name;

    public Integer getCode() {
        return code;
    }

    public void setCode(Integer code) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    Aggregation(Integer code, String name) {
        this.code = code;
        this.name = name;
    }
}
