package org.example.exceldownload;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.example.exceldownload.excel.service.ExcelUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
public class IndexController {


    @GetMapping("/exceldownload")
    public void index(HttpServletRequest request, HttpServletResponse response) throws IOException {

        /* 예시 데이터 */
        List<ExampleExcelDto> exampleExcelDtos = new ArrayList<>();
        exampleExcelDtos.add(new ExampleExcelDto("ldw", "19940328", "M", "이동욱"));
        exampleExcelDtos.add(new ExampleExcelDto("ldw2", "19950328", "W", "강동욱"));
        exampleExcelDtos.add(new ExampleExcelDto("ldw3", "19960328", "M", "우동욱"));

        // 엑셀다운로드
        ExcelUtils.downloadExcelFile(response, exampleExcelDtos , ExampleExcelDto.class);
    }

}
