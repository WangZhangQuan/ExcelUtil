package com.cat.excel.util.test;

import com.cat.excel.util.ExcelExportUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.*;

public class Test {

    @org.junit.Test
    public void t1() throws IOException, InvalidFormatException {
        // 准备model
        Map<String, Object> model = new HashMap<String, Object>();
        List<String> sheets = new ArrayList<String>();
        sheets.add("测试一");
        sheets.add("测试二");
        sheets.add("测试三");

        Random random = new Random(System.currentTimeMillis());

        for (String sheet : sheets) {
            Map<String, Object> sheetModel = new HashMap<String, Object>();
            sheetModel.put("姓名", "Cat_" + random.nextInt(999));
            sheetModel.put("年龄", random.nextInt(30));
            sheetModel.put("性别", (random.nextInt() << 15) == 1 ? '男' : '女');
            model.put(sheet, sheetModel);
        }

        model.put("测试", sheets);

        // 创建excel工具
        ExcelExportUtil instance = ExcelExportUtil.createInstance(model, Test.class.getResource("/").getPath() + "test.xlsx");
        instance.parse_0_1();
        instance.writeBook("d:/test.xlsx");
    }

}
