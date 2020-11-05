package com;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReadStatus;
import org.jxls.reader.XLSReader;
import org.springframework.boot.test.context.SpringBootTest;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootTest
class JxlsApplicationTests {

    @Test
    void contextLoads() throws IOException, SAXException, InvalidFormatException {
        //需要导出的参数的xls文件路径
        String path = "E:\\Project\\JXLS\\src\\main\\resources\\template\\departmentdata.xls";
        //配置文件路径
        String xmlPath = "E:\\Project\\JXLS\\src\\main\\resources\\template\\example.xml";
        InputStream inputXML = new BufferedInputStream(new FileInputStream(xmlPath));
        InputStream inputXLS = new BufferedInputStream(new FileInputStream(new File(path)));
        XLSReader mainReader = ReaderBuilder.buildFromXML(inputXML);
        //这个参数对应xml配置文件的参数
        List<Employee> employeeList= new ArrayList<>();
        Map<String, Object> beans = new HashMap<>();
        beans.put("employees", employeeList);
        //read是读取数据
        XLSReadStatus readStatus = mainReader.read(inputXLS, beans);
        //是否读取成功
        boolean statusOK = readStatus.isStatusOK();
        System.out.println("读取结果：" + statusOK);
        //测试结果
        for (Employee employee : employeeList) {
            System.out.println(employee);
        }
    }

}
