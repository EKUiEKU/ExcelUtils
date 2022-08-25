package top.xizia.utils.example.controller;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import top.xizia.utils.example.entity.FlightNoticeArrivalDetailDTO;
import top.xizia.utils.example.entity.FlightNoticeArrivalReportUnit;
import top.xizia.utils.example.entity.Fourth;
import top.xizia.utils.example.entity.Third;
import top.xizia.utils.poi.ExcelUtils;
import top.xizia.utils.poi.TemplateStyleListener;

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

        ExcelUtils.downloadExcel(response, ret, "二级大标题", "一级大标题");
    }


    @GetMapping("/download2")
    public void test2(HttpServletResponse response) throws IOException, IllegalAccessException, InstantiationException {
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

        ExcelUtils.downloadExcel(response, ret, "二级大标题", "一级大标题", new TemplateStyleListener() {
            @Override
            public List<List<List<Object>>> onAddHeadInfoEvent() {
                List<List<List<Object>>> list = new ArrayList<>();

                /**
                 * 第一行
                 */
                List<List<Object>> row1 = new ArrayList<>();
                row1.add(Arrays.asList(1, 1, "航班号"));
                row1.add(Arrays.asList(2, 1, "CZ1234"));
                row1.add(Arrays.asList(5, 1, "航班日期:"));
                row1.add(Arrays.asList(6, 1, "2022-05-13"));

                list.add(row1);


                /**
                 * 第二行
                 */
                List<List<Object>> row2 = new ArrayList<>();
                row2.add(Arrays.asList(1, 1, "起始站: 北京"));
                row2.add(Arrays.asList(5, 1, "到达站: 上海"));

                list.add(row2);


                return list;
            }

            @Override
            public List<List<List<Object>>> onAddBottomInfoEvent(int left, int right, int top, int bottom) {
                List<List<List<Object>>> list = new ArrayList<>();

                /**
                 * 第一行
                 */
                List<List<Object>> row1 = new ArrayList<>();
                row1.add(Arrays.asList(1, 1, "制表人"));
                row1.add(Arrays.asList(2, 1, "吴多聪"));
                list.add(row1);


                /**
                 * 第二行
                 */
                List<List<Object>> row2 = new ArrayList<>();
                row2.add(Arrays.asList(1, 1, "制表时间"));
                row2.add(Arrays.asList(2, 1, "2022-8-23"));

                list.add(row2);


                return list;
            }

            @Override
            public void onModifyStyleEvent(ExcelWriter writer, int left, int right, int top, int bottom) {

            }
        });
    }
}
