package me.mhlee.excelport.annotation;

import me.mhlee.excelport.Excelport;
import me.mhlee.excelport.annotation.dto.Member;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;

import org.apache.commons.lang3.RandomUtils;
import org.apache.commons.lang3.time.StopWatch;
import org.junit.jupiter.api.Test;

public class PerformanceTest {
    public static final long MAX_ROW = 200_000;

    @Test
    public void performanceTest() throws FileNotFoundException {
        StopWatch stopWatch = new StopWatch();
        List<Member> members = new LinkedList<>();

        for(long ii =0; ii < MAX_ROW; ii++) {
            members.add(Member.create(UUID.randomUUID().toString(), (RandomUtils.nextInt() % 80) + 1));
        }

        System.out.println("테스트 객체 생성 완료!!");

        String tempDir = System.getProperty("java.io.tmpdir");
        System.out.println(tempDir);


        OutputStream excelOs = new FileOutputStream(new File(tempDir, "performance_out.xlsx"));
        stopWatch.start();
        Excelport.toExcel(excelOs, members.iterator());
        stopWatch.stop();
        System.out.println("time :: " + stopWatch.getTime());

    }
}
