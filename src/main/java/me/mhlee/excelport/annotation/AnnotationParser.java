package me.mhlee.excelport.annotation;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;

public class AnnotationParser {

    /** 객체에 선언된 Excel annotation 정보와 필드명을 order 순서에 맞게 정렬하여 리턴
     * @param template Excel annotation 정보가 선언된 String array
     * @return
     */
    public static List<ExcelField> extractExcelColumnsFromString(String[] template) {
        List<ExcelField> result = new ArrayList<>();

        for(String item : template) {
            ExcelField excelField = new ExcelField();
            String[] fields = item.split(",");

            for (String field: fields) {
                excelField.setFieldByStringTemplate(field);
            }

            if (Objects.isNull(excelField.getFieldName())) {
                System.out.println("fieldName is null!! ==> " + excelField);
            }

            System.out.println(excelField);

            result.add(excelField);
        }

        return result.stream()
                .sorted()
                .collect(Collectors.toList());
    }

    /** 객체에 선언된 Excel annotation 정보와 필드명을 order 순서에 맞게 정렬하여 리턴
     * @param obj Excel annotation 이 선언된 object
     * @return
     */
    public static List<ExcelField> extractExcelColumns(Object obj) {
        Field[] fields = obj.getClass().getDeclaredFields();

        return Arrays.stream(fields)
                .map(it -> Pair.of(it, it.getDeclaredAnnotations()))
                .map(it -> getExcelAnnotationIfExits(it.getLeft(), it.getRight()))
                .filter(it -> Objects.nonNull(it))
                .sorted()
                .collect(Collectors.toList());
    }

    /** 필드명과 Excel annotation 정보를 조합하여 CellInfo 클래스로 리턴
     * @param f
     * @param annotations
     * @return
     */
    private static ExcelField getExcelAnnotationIfExits(Field f, Annotation[] annotations) {
        return Arrays.stream(annotations)
                .filter(it -> it instanceof Excel)
                .map(it -> (Excel)it)
                .findFirst()
                .map(it -> ExcelField.create(f.getName(), it.name(), it.order(), it.align(), it.format()))
                .orElse(null);
    }

    /** DTO(or VO) 형태의 객체를 Map<필드명, 값> 형태의 Map 으로 리턴
     * @param obj
     * @return
     */
    private static Map<String, Object> toMap(Object obj) {
        Field[] fields = obj.getClass().getDeclaredFields();

        Map<String, Object> kv = new HashMap<>();
        for(Field f : fields) {
            Object value = null;
            try {
                Method method = obj.getClass().getMethod("get" + StringUtils.capitalize(f.getName()));
                value = method.invoke(obj);
            } catch (NoSuchMethodException | InvocationTargetException | IllegalAccessException e) {
                f.setAccessible(true);
                try {
                    value = f.get(obj);
                } catch (IllegalAccessException ex) {
                    ex.printStackTrace();
                }
            } finally {
                kv.put(f.getName(), value);
            }
        }

        return kv;
    }

    /** 헤더명만 list 로 출력
     * @param columns
     * @return
     */
    private static List<String> extractHeader(List<ExcelField> columns) {
        return columns.stream()
                .map(it -> it.getHeader())
                .collect(Collectors.toList());
    }

    /** 입력된 Object 에서 출력값(date value)을 order 순서에 맞게 List로 변환
     * columns 에는 Excel annotation 의 메타정보 수록
     * @param columns
     * @param obj
     * @return
     */
    private static List<Object> extractRow(List<ExcelField> columns, Object obj) {
        if (obj instanceof Map)
            return extractRow(columns, (Map)obj);
        else
            return extractRow(columns, toMap(obj));
    }

    private static List<Object> extractRow(List<ExcelField> columns, Map<String, Object> map) {
        return columns.stream()
                .map(it -> map.get(it.getFieldName()))
                .collect(Collectors.toList());
    }

    /** Excel annotation 메타정보에서 헤더컬럼만 추출
     * @param columns
     * @return
     */
    public static List<String> toHeader(List<ExcelField> columns) {
        return extractHeader(columns);
    }


    /** 입력된 Object 의 data 를 order 순서에 맞게 List 로 변환
     * @param obj
     * @return
     */
    public static List<Object> toRow(Object obj, List<ExcelField> parsedExcels) {
        return extractRow(parsedExcels, obj);
    }

}

