package me.mhlee.excelport.annotation;

import me.mhlee.excelport.Excelport;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.RandomUtils;
import org.junit.jupiter.api.Test;

public class ExcelCreateTest {

    public static class Member {
        @Excel
        String name;

        @Excel
        int age;

        @Excel
        long point;

        @Excel
        float score;

        @Excel
        LocalDate birthday;

        @Excel
        LocalDateTime joinedAt;

        @Excel
        Date lastAccess;

        @Excel
        List<Integer> numbers;

        public String getPoint() {
            if (point > 80) {
                return "Gold member";
            } else if (point > 50) {
                return "Silver member";
            } else {
                return "bronze member";
            }
        }

        public Member(String name, int age) {
            this.name = name;
            this.age = age;

            this.point = RandomUtils.nextLong() % 100;
            this.score = RandomUtils.nextFloat() % 100f;
            this.joinedAt = LocalDateTime.now().minusDays(RandomUtils.nextInt() % 365);
            this.birthday = LocalDate.now().minusDays(RandomUtils.nextInt() % (365 * age));
            this.lastAccess = new Date();

            this.numbers = new ArrayList<>();

            for(int ii = 0; ii < 10; ii++) {
                numbers.add(RandomUtils.nextInt());
            }
        }

        public static Member create(String name, int age) {
            return new Member(name, age);
        }
    }

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
