package com.beiming.doexcel.controller;

import com.beiming.doexcel.service.DoExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RestController;

import java.sql.SQLException;

@RestController
public class DoExcelController {

    @Autowired
    DoExcelService doExcelService;

    @Value("${pathWay}")
    String pathWay;

    @GetMapping("/insert/{filePath}/{tableName}/{startNum}/{endNum}")
    public String sayHello(@PathVariable String filePath, @PathVariable String tableName,
                            @PathVariable String startNum, @PathVariable String endNum) throws SQLException, ClassNotFoundException {

        StringBuffer message = doExcelService.doSave(pathWay + filePath + "/",tableName,Integer.parseInt(startNum),Integer.parseInt(endNum));
        return message.append("<br>filePath=" + filePath + "<br>tableName=" + tableName + "<br>startNum=" + startNum + "<br>endNum=" + endNum).toString();
    }
}
