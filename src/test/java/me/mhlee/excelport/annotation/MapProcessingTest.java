package me.mhlee.excelport.annotation;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.junit.jupiter.api.Test;


public class MapProcessingTest {

    public static class Member {
        @Excel
        String name;

        @Excel
        String age;
    }

    @Test
    public void mapToDataClass() {
        List<Map<String, Object>> list = new ArrayList<>();

        Map<String, Object> row1 = new HashMap<>();
        row1.put("name", "SH,Lee");;
        row1.put("age", "9");

        Map<String, Object> row2 = new HashMap<>();
        row2.put("name", "SJ,Lee");;
        row2.put("age", "4");

        list.add(row1);
        list.add(row2);

        Member m = toDataClass(row1, Member.class);
    }


    public <T> T toDataClass(Map<String, Object> map, Class<T> clazz) {
        try {
            Object obj = clazz.newInstance();
            Field[] fields = obj.getClass().getDeclaredFields();

            Set<String> keys = map.keySet();

            for(String key : keys) {
                System.out.println(key);
            }

            System.out.println("============================");

            for(Field f : fields) {
                System.out.println(f.getName());
            }
            return null;
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
            throw new RuntimeException(e.getMessage());
        }
    }
}
