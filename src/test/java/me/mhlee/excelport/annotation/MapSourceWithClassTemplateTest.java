package me.mhlee.excelport.annotation;

import me.mhlee.excelport.Excelport;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;


public class MapSourceWithClassTemplateTest {

    public static class Member {
        @Excel(order = 1)
        String name;

        @Excel
        int age;

        @Excel(name = "별명", order = 2)
        String nick_name;

        @Excel(name = "포인트")
        long my_point;
    }

    @Test
    public void mapToDataClass() throws FileNotFoundException {
        List<Map<String, Object>> list = new ArrayList<>();

        Map<String, Object> row1 = new HashMap<>();
        row1.put("name", "SH,Lee");;
        row1.put("age", 9);
        row1.put("nick_name", "hello");
        row1.put("my_point", 2.2);

        Map<String, Object> row2 = new HashMap<>();
        row2.put("name", "SJ,Lee");;
        row2.put("age", 4);
        row2.put("nick_name", "world");
        row2.put("my_point", 1.1);

        list.add(row1);
        list.add(row2);

        String tempDir = System.getProperty("java.io.tmpdir");
        System.out.println(tempDir);

        // excel 출력
        OutputStream excelOs = new FileOutputStream(new File(tempDir, "out_with_map_class_template.xlsx"));
        Excelport.toExcel(excelOs, list.iterator(), Member.class);
    }

}
