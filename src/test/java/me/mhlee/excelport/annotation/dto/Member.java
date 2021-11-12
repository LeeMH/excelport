package me.mhlee.excelport.annotation.dto;

import me.mhlee.excelport.annotation.Excel;
import me.mhlee.excelport.annotation.ExcelCreateTest;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.RandomUtils;

public class Member {
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
