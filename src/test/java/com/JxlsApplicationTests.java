package com;

import com.entity.Department;
import com.entity.Employee;
import com.util.DateCommand;
import com.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReadStatus;
import org.jxls.reader.XLSReader;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.springframework.boot.test.context.SpringBootTest;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.*;

@SpringBootTest
@Slf4j
class JxlsApplicationTests {

    String location = System.getProperty("user.dir") + "\\src\\main\\resources\\template\\";

    @Test
    Map<String, Object> readExcel() throws IOException, SAXException, InvalidFormatException {
        //需要导出的参数的xls文件路径
        String path = location + "departmentdata.xls";
        //配置文件路径
        String xmlPath = location + "example.xml";

        InputStream inputXML = new BufferedInputStream(new FileInputStream(xmlPath));
        InputStream inputXLS = new BufferedInputStream(new FileInputStream(new File(path)));
        XLSReader mainReader = ReaderBuilder.buildFromXML(inputXML);

        // 存放excel数据 这个参数对应xml配置文件的参数
        List<Employee> employeeList= new ArrayList<>();
        List<Employee> employeeList2= new ArrayList<>();
        List<Department> departments = new ArrayList();
        Map<String, Object> beans = new HashMap<>();
        beans.put("employees", employeeList);
        beans.put("employees2", employeeList2);
        beans.put("departments", departments);

        //读取数据
        XLSReadStatus readStatus = mainReader.read(inputXLS, beans);

        //是否读取成功
        boolean statusOK = readStatus.isStatusOK();
        System.out.println("读取结果：" + statusOK);
        //测试结果
        for (Employee employee : employeeList) {
            System.out.println(employee);
        }
        for (Department department:departments){
            System.out.println(department);
        }
        // jxls会将读取的数组和最后一条数据写入beans(模板文件<loop>中的items、var)
        System.out.println("beans:"+beans);
        return beans;
    }

    @Test
    void writeExcel() throws IOException, SAXException, InvalidFormatException {

        Map<String, Object> model = readExcel();
        String templateSum = location + "writeTemplate.xlsx";
        String exportFilePath = location + "exportExcel.xlsx";

        try ( InputStream is =  new BufferedInputStream(new FileInputStream(new File(templateSum)));) {
            try (OutputStream os = new FileOutputStream(exportFilePath);) {
                Context context = new Context();
                if (model != null) {
                    for (String key : model.keySet()) {
                        context.putVar(key, model.get(key));
                    }
                }
                JxlsHelper jxlsHelper = JxlsHelper.getInstance();

                //第一种方式
                Transformer transformer = PoiTransformer.createTransformer(is, os);
                // 执行excel表达式，如果没有这个语句，则excel中的表达式不起作用
                transformer.setEvaluateFormulas(true);

                //自定义功能
                //jxls2.6以上
                //获得配置
                JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig()
                        .getExpressionEvaluator();
                //函数强制，自定义功能
                Map<String, Object> funcs = new HashMap<String, Object>();
                funcs.put("utils", new ExcelUtil()); //添加自定义功能

                JexlBuilder jb = new JexlBuilder();
                jb.namespaces(funcs);
                //jb.silent(true); //设置静默模式，不报警告
                JexlEngine je = jb.create();
                evaluator.setJexlEngine(je);

                //jxls2.6以下
//                //获得配置
//                JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig()
//                        .getExpressionEvaluator();
//                //设置静默模式，不报警告
//                //evaluator.getJexlEngine().setSilent(true);
//                //函数强制，自定义功能
//                Map<String, Object> funcs = new HashMap<String, Object>();
//                funcs.put("utils", new JxlsUtils()); //添加自定义功能
//                evaluator.getJexlEngine().setFunctions(funcs);

                XlsCommentAreaBuilder.addCommandMapping(DateCommand.COMMAND_NAME,DateCommand.class);
                ((PoiTransformer) transformer).getWorkbook().setForceFormulaRecalculation(true);

                jxlsHelper.processTemplate(context, transformer);

                //第二种方式
//                JxlsHelper jxlsHelper = JxlsHelper.getInstance();
//                jxlsHelper.setEvaluateFormulas(true);
//                jxlsHelper.processTemplate(is,os,context);
            } catch (IOException e) {
                log.error("导出位置异常，请检查！！,{}", exportFilePath);
                throw e;
            }
        } catch (IOException e) {
            log.error("模板文件异常，请检查！！,{}", templateSum);
            throw e;
        }
    }

    @Test
    void gridTest() throws InvalidFormatException, SAXException, IOException {
        Map<String, Object> model = readExcel();
        String templateSum = location + "grid_template.xls";
        String exportFilePath = location + "exportExcel.xlsx";

        try ( InputStream is =  new BufferedInputStream(new FileInputStream(new File(templateSum)));) {
            try (OutputStream os = new FileOutputStream(exportFilePath);) {
                Context context = new Context();
                if (model != null) {
                    for (String key : model.keySet()) {
                        context.putVar(key, model.get(key));
                    }
                }
                // 列名
                context.putVar("headers", Arrays.asList("Name", "Birthday", "Payment"));
                context.putVar("data", model.get("employees"));
                // 对应beans中的属性
                JxlsHelper.getInstance().processGridTemplateAtCell(is, os, context, "name,birthDate,payment", "Sheet1!A1");
            } catch (IOException e) {
                log.error("导出位置异常，请检查！！,{}", exportFilePath);
                throw e;
            }
        } catch (IOException e) {
            log.error("模板文件异常，请检查！！,{}", templateSum);
            throw e;
        }
    }
}
