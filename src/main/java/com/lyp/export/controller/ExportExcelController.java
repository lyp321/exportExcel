package com.lyp.export.controller;


import com.lyp.export.util.ExcelUtils;
import com.lyp.export.util.Person;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.time.DateUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @program: exportexcel
 * @description: 导出excel控制层
 * @author: Liyp
 * @create: 2019-11-11 17:01
 **/
@RestController
@Slf4j
public class ExportExcelController {
    @RequestMapping("export")
    public void export(HttpServletResponse response){

        //模拟从数据库获取需要导出的数据
        List<Person> personList = new ArrayList<>();
        Person person1 = new Person("路飞","1",new Date());
        Person person2 = new Person("娜美","2", DateUtils.addDays(new Date(),3));
        Person person3 = new Person("索隆","1", DateUtils.addDays(new Date(),10));
        Person person4 = new Person("小狸猫","1", DateUtils.addDays(new Date(),-10));
        personList.add(person1);
        personList.add(person2);
        personList.add(person3);
        personList.add(person4);

        //导出操作
       ExcelUtils.exportExcel(personList,"花名册","草帽一伙",Person.class,"海贼王.xls",response);
    }

    @RequestMapping("importExcel")
    public void importExcel(){
        String filePath = "F:\\海贼王.xls";
        //解析excel，
        List<Person> personList = ExcelUtils.importExcel(filePath,1,1,Person.class);
        //也可以使用MultipartFile,使用 FileUtil.importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class<T> pojoClass)导入
        System.out.println("导入数据一共【"+personList.size()+"】行");

        //TODO 保存数据库
    }
}