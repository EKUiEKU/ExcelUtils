package top.xizia.utils.example.entity;


import top.xizia.utils.poi.Aggregation;
import top.xizia.utils.poi.Excel;

import java.math.BigDecimal;

/**
 * @author: WSC
 * @DATE: 2022/8/1
 * @DESCRIBE: 进港播报统计的详情结果
 **/
public class FlightNoticeArrivalDetailDTO{
    /**
     * 序号
     */
    @Excel(value = "序号", isIndex = true, width = 10)
    private Long idx;
    /**
     * 承运人
     */
    @Excel(value = "承运人", sort = 1, width = 15)
    private String carrier;
    /**
     * 件数
     */
    @Excel(value = "件数", aggregation = Aggregation.SUM, sort = 2, width = 10)
    private Integer pieces;

    @Excel(value = "重量", aggregation = Aggregation.SUM, sort = 2, width = 10)
    private BigDecimal weight;
    /**
     * 航班日期
     */
    @Excel(value = "航班日期", sort = 2, width = 15)
    private Long flightDate;
    /**
     * 航班号
     */
    @Excel(value = "航班号", sort = 3, width = 15)
    private String flightNo;
    /**
     * 航班始发站
     */
    @Excel(value = "航班始发站", sort = 4, width = 10)
    private String departure;
    /**
     * 航班目的站
     */
    @Excel(value = "航班目的站", sort = 5, width = 10)
    private String arrival;
    /**
     * 货单数 计划运单票数
     */
    @Excel(value = "货单数", sort = 6, width = 10)
    private Integer waybillNumbers;
    /**
     * 货物件数
     */
    @Excel(value = "货物件数", sort = 7, width = 10)
    private Integer cargoPieces;
    /**
     * 货物重量
     */
    @Excel(value = "货物重量", sort = 8, width = 10)
    private BigDecimal cargoWeight;
    /**
     * 邮件重量
     */
    @Excel(value = "邮件重量", sort = 9, width = 10)
    private BigDecimal mailWeight;
    /**
     * 货邮总重量
     */
    @Excel(value = "货邮总重量", sort = 10, width = 10)
    private BigDecimal totalWeight;
    /**
     * 国内中转货量
     */
    @Excel(value = "国内中转货量", sort = 11, width = 10)
    private BigDecimal domesticTransitWeight;
    /**
     * 国际中转货量
     */
    @Excel(value = "国际中转货量", sort = 12, width = 10)
    private BigDecimal internationalTransitWeight;
    /**
     * 多表头 分拣货邮
     */
    @Excel(value = "分拣货邮", sort = 99, isMultipleHeaders = true, width = 10)
    private FlightNoticeArrivalReportUnit sortingMail;
    /**
     * 多表头 交卸交接
     */
    @Excel(value = "交卸交接", sort = 100, isMultipleHeaders = true, width = 10)
    private FlightNoticeArrivalReportUnit handover;
    /**
     * 多表头 交接袋
     */
    @Excel(value = "交接袋", sort = 101, isMultipleHeaders = true, width = 10)
    private FlightNoticeArrivalReportUnit businessBag;
    /**
     * 多表头 提单交货时间差
     */
    @Excel(value = "提单交货时间差", sort = 102, isMultipleHeaders = true, width = 10)
    private FlightNoticeArrivalReportUnit orderDeliveryTimeDifference;

    public Long getIdx() {
        return idx;
    }

    public void setIdx(Long idx) {
        this.idx = idx;
    }

    public String getCarrier() {
        return carrier;
    }

    public void setCarrier(String carrier) {
        this.carrier = carrier;
    }

    public Long getFlightDate() {
        return flightDate;
    }

    public void setFlightDate(Long flightDate) {
        this.flightDate = flightDate;
    }

    public String getFlightNo() {
        return flightNo;
    }

    public void setFlightNo(String flightNo) {
        this.flightNo = flightNo;
    }

    public String getDeparture() {
        return departure;
    }

    public void setDeparture(String departure) {
        this.departure = departure;
    }

    public String getArrival() {
        return arrival;
    }

    public void setArrival(String arrival) {
        this.arrival = arrival;
    }

    public Integer getWaybillNumbers() {
        return waybillNumbers;
    }

    public void setWaybillNumbers(Integer waybillNumbers) {
        this.waybillNumbers = waybillNumbers;
    }

    public Integer getCargoPieces() {
        return cargoPieces;
    }

    public void setCargoPieces(Integer cargoPieces) {
        this.cargoPieces = cargoPieces;
    }

    public BigDecimal getCargoWeight() {
        return cargoWeight;
    }

    public void setCargoWeight(BigDecimal cargoWeight) {
        this.cargoWeight = cargoWeight;
    }

    public BigDecimal getMailWeight() {
        return mailWeight;
    }

    public void setMailWeight(BigDecimal mailWeight) {
        this.mailWeight = mailWeight;
    }

    public BigDecimal getTotalWeight() {
        return totalWeight;
    }

    public void setTotalWeight(BigDecimal totalWeight) {
        this.totalWeight = totalWeight;
    }

    public BigDecimal getDomesticTransitWeight() {
        return domesticTransitWeight;
    }

    public void setDomesticTransitWeight(BigDecimal domesticTransitWeight) {
        this.domesticTransitWeight = domesticTransitWeight;
    }

    public BigDecimal getInternationalTransitWeight() {
        return internationalTransitWeight;
    }

    public void setInternationalTransitWeight(BigDecimal internationalTransitWeight) {
        this.internationalTransitWeight = internationalTransitWeight;
    }

    public FlightNoticeArrivalReportUnit getSortingMail() {
        return sortingMail;
    }

    public void setSortingMail(FlightNoticeArrivalReportUnit sortingMail) {
        this.sortingMail = sortingMail;
    }

    public FlightNoticeArrivalReportUnit getHandover() {
        return handover;
    }

    public void setHandover(FlightNoticeArrivalReportUnit handover) {
        this.handover = handover;
    }

    public FlightNoticeArrivalReportUnit getBusinessBag() {
        return businessBag;
    }

    public void setBusinessBag(FlightNoticeArrivalReportUnit businessBag) {
        this.businessBag = businessBag;
    }

    public FlightNoticeArrivalReportUnit getOrderDeliveryTimeDifference() {
        return orderDeliveryTimeDifference;
    }

    public void setOrderDeliveryTimeDifference(FlightNoticeArrivalReportUnit orderDeliveryTimeDifference) {
        this.orderDeliveryTimeDifference = orderDeliveryTimeDifference;
    }

    public Integer getPieces() {
        return pieces;
    }

    public void setPieces(Integer pieces) {
        this.pieces = pieces;
    }
}
