# excelport

## What is excelport ?

---

Excelport is library to help excel export easily.

## install

---

- add repository your pom.xml
```
    <repositories>
        <repository>
            <id>excelport-snapshot</id>
            <url>https://github.com/LeeMH/maven-repo/tree/master/snapshots/me/mhlee/excelport/1.0-SNAPSHOT</url>
        </repository>
    </repositories>

```

- add dependency
```
        <dependency>
            <groupId>me.mhlee</groupId>
            <artifactId>excelport</artifactId>
            <version>1.0-SNAPSHOT</version>
        </dependency>
```

## how to use

---
 
### 1. add @excel annotation your DTO or VO class

- name : excel column header. default value is variable name.
- order : column order. default value is 999. 
- align : align option. default value is None. (number type variable right align)
- value : if variable has getter method, fill column as getter method result

```
        @Excel(name = "age", order = 2, align = Align.RIGHT)
        int age;
```

### 2. export excel only 1 line
```
        OutputStream os = new FileOutputStream("out.xlsx");
        ExcelService.toExcel(os, members.iterator());
```

### 3. example
```
public class App {

    public static class Member {
        @Excel(name = "member name", order = 1)
        String name;

        @Excel(name = "age", order = 2, align = Align.RIGHT)
        int age;

        @Excel(name = "join date")
        LocalDate joinedAt = LocalDate.now();

        @Excel(name = "member point", order = 2, align = Align.LEFT)
        long point = RandomUtils.nextLong() % 100;

        long excludeField = RandomUtils.nextLong();

        public Member(String name, int age) {
            this.name = name;
            this.age = age;
        }

        public String getPoint() {
            String result = null;

            if (point > 90) result = "over 90 point";
            else if (point > 50) result = "over 50 point";
            else result = "under 50 point";

            return result;
        }
    }

    public static void main(String...args) throws FileNotFoundException {
        List<Member> members = new ArrayList<>();

        Member member1 = new Member("SeoHo, Lee", 9);
        Member member2 = new Member("SeoJung, Lee", 4);
        members.add(member1);
        members.add(member2);

        OutputStream os = new FileOutputStream("out.xlsx");
        ExcelService.toExcel(os, members.iterator());
    }
}
```

### 4. result

![excelport-result](https://user-images.githubusercontent.com/5078531/141324070-4b3fc604-eb38-4fa9-9ef3-3736f0ecd267.png)


## additional information

---

### available export csv, json

- use toJson(OutputStream, Iterator) 

- use toCsv(OutputStream, Iterator)

### for my-batis user 

- If you handle large size data, use  Cursor instead of List to avoid OOM(Out Of Memory)
- List, Cursor implement Iterator interface
- Cursor should be under transaction!! So, if you want to use Cursor first start transaction and don't finish transaction until use it.
``` 
List<Member> selectSuperLargeDataSetAsList(Map<String, Object> params);
 
Cursor<Member> selectSuperLargeDataSetAsCursor(Map<String, Object> params);
```

### for web service

- You can send excel using OutputStream of HttpServletResponse.
```
@RequestMapping(value = "/download", method = RequestMethod.GET)
public void download(HttpServletResponse response) {
    ... something do your logic
    ... set http resonse header
    
    ExcelService.toExcel(response.getOutputStream(),members);
    
    ...
    ...
}
```
