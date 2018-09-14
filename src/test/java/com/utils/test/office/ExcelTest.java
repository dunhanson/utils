package com.utils.test.office;

import com.utils.office.ExcelUtils;
import org.apache.commons.lang3.StringUtils;
import org.junit.After;
import org.junit.Test;

import java.util.*;

public class ExcelTest {

    @After
    public void after() {
        System.out.println("finish...");
    }

    @Test
    public void read() {
        String filePath = "D:\\【修订】检查投标文件（通用版）.xlsx";
        List<Object[]> result = ExcelUtils.readWorkbookByLimit(filePath, 3, 0);

        List<Object> models = new ArrayList<>();
        Map<Integer, List<Object>> modules = new HashMap<>();
        Map<Integer, List<Object>> rules = new HashMap<>();
        for (Object[] arr: result) {
            if(isModel(arr)) {//Modle
                models.add(arr[0]);
            } else if(isModule(arr)) {//Module
                //添加Module
                int hashCode = models.get(models.size() - 1).hashCode();
                if(modules.containsKey(hashCode)) {
                    modules.get(hashCode).add(arr[1]);
                } else {
                    List<Object> module = new ArrayList<>();
                    module.add(arr[1]);
                    modules.put(hashCode, module);
                }
                //添加Rule
                List<Object> rule = new ArrayList<>();
                rule.add(arr[2]);
                rules.put(arr[1].hashCode(), rule);
            } else if(isRule(arr)) {//Rule
                int hashCode = models.get(models.size() - 1).hashCode();
                List<Object> module = modules.get(hashCode);
                rules.get(module.get(module.size() - 1).hashCode()).add(arr[2]);
            }
        }

        models.forEach(obj->{
            System.out.println(obj);
            List<Object> module = modules.get(obj.hashCode());
            module.forEach(obj1->{
                System.out.println(" " + (module.indexOf(obj1) + 1) + " " + obj1);
                rules.get(obj1.hashCode()).forEach(obj2->{
                    System.out.println("  "+obj2);
                });
            });
            System.out.println();
        });
    }

    public boolean isModel(Object[] arr) {
        toStrArr(arr);
        if(StringUtils.isNotBlank((String)arr[0]) && StringUtils.isBlank((String)arr[1])
                && StringUtils.isBlank((String)arr[2])) {
            return true;
        }
        return false;
    }

    public boolean isModule(Object[] arr) {
        toStrArr(arr);
        if(StringUtils.isNotBlank((String)arr[0]) && StringUtils.isNotBlank((String)arr[1])
                && StringUtils.isNotBlank((String)arr[2])) {
            return true;
        }
        return false;
    }

    public boolean isRule(Object[] arr) {
        toStrArr(arr);
        if(StringUtils.isBlank((String)arr[0]) && StringUtils.isBlank((String)arr[1])
                && StringUtils.isNotBlank((String)arr[2])) {
            return true;
        }
        return false;
    }

    public void toStrArr(Object[] arr) {
        arr[0] = arr[0] == null ? "" : String.valueOf(arr[0]);
        arr[1] = arr[1] == null ? "" : String.valueOf(arr[1]);
        arr[2] = arr[2] == null ? "" : String.valueOf(arr[2]);
    }


}
