package top.xizia.utils.example.controller;

import cn.hutool.poi.excel.ExcelUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import top.xizia.utils.example.entity.FlightNoticeArrivalDetailDTO;
import top.xizia.utils.example.entity.FlightNoticeArrivalReportUnit;
import top.xizia.utils.example.entity.Fourth;
import top.xizia.utils.example.entity.Third;
import top.xizia.utils.poi.ExcelUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author: WSC
 * @DATE: 2022/8/10
 * @DESCRIBE:
 **/
@RestController
public class ExampleController {
    @GetMapping("/download")
    public void test(HttpServletResponse response) throws IOException, IllegalAccessException, InstantiationException {
        FlightNoticeArrivalDetailDTO dto = new FlightNoticeArrivalDetailDTO();
        dto.setPieces(2);
        Fourth fourth = new Fourth();
        fourth.setTest3("测试3");
        fourth.setTest4("测试4");
        fourth.setTest5("测试5");


        Third third = new Third();
        third.setTest1("测试1");
        third.setTest2("测试2");
        third.setFourth(fourth);

        FlightNoticeArrivalReportUnit unit = new FlightNoticeArrivalReportUnit();
        unit.setDateLine(0L);
        unit.setOvertimeTime(0L);
        unit.setSpendTime(0L);
        unit.setThird(third);

        dto.setArrival("CAN");
        dto.setBusinessBag(unit);
        dto.setHandover(unit);
        dto.setSortingMail(unit);
        dto.setOrderDeliveryTimeDifference(unit);

        List<FlightNoticeArrivalDetailDTO> ret = new ArrayList<>();

        /**
         * 不建议导出10W条数据以上,如想尝试则需调整JVM的堆上限
         */
        for (int i = 0; i < 1000; i++) {
            dto.setFlightNo(System.currentTimeMillis() + "");
            ret.add(dto);
        }

        ExcelUtils.downloadExcel(response, ret);
    }
}
