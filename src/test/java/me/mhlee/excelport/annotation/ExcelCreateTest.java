package me.mhlee.excelport.annotation;

import me.mhlee.excelport.Excelport;
import me.mhlee.excelport.annotation.dto.Member;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

public class ExcelCreateTest {

    @Test
    public void createExcelFile() throws FileNotFoundException {
        String tempDir = System.getProperty("java.io.tmpdir");
        System.out.println(tempDir);

        List<Member> members = new ArrayList<>();
        members.add(Member.create("SH, LEE", 9));
        members.add(Member.create("SJ, LEE", 4));

        // excel 출력
        OutputStream excelOs = new FileOutputStream(new File(tempDir, "out.xlsx"));
        Excelport.toExcel(excelOs, members.iterator());

        // csv 출력
        OutputStream csvOs = new FileOutputStream(new File(tempDir, "out.csv"));
        Excelport.toCsv(csvOs, members.iterator());

        // json 출력
        OutputStream jsonOs = new FileOutputStream(new File(tempDir, "out.json.txt"));
        Excelport.toJson(jsonOs, members.iterator());
    }
}
