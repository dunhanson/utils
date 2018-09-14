package com.utils.test.office;

import com.utils.office.ExcelUtils;
import org.junit.After;
import org.junit.Test;
import java.util.Arrays;

public class ExcelTest {

    @After
    public void after() {
        System.out.println("finish...");
    }

    @Test
    public void read() {
        String filePath = "D:\\【修订】检查投标文件（通用版）.xlsx";
        ExcelUtils.readWorkbook(filePath).forEach(obj->{
            System.out.println(Arrays.toString(obj));
        });
    }



}
