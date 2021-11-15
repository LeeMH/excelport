# excelport

## 1. excelport ?

---

excelport 는 Excel + Export 의 합성어 입니다.

excelport 는 단 1줄의 코드로 Collection 타입의 데이터를 엑셀로 변환 합니다. 

## 2. 설치 방법

---

- pom.xml 파일에 repository 정보 추가
```xml
    <repositories>
        <repository>
            <id>excelport-snapshot</id>
            <url>https://github.com/LeeMH/maven-repo/tree/master/snapshots/me/mhlee/excelport/1.1-SNAPSHOT</url>
        </repository>
    </repositories>

```

- 의존성 추가
```xml
        <dependency>
            <groupId>me.mhlee</groupId>
            <artifactId>excelport</artifactId>
            <version>1.1-SNAPSHOT</version>
        </dependency>
```

## 3. 사용방법

---

단 1줄의 toExcel 메소드 호출로 엑셀 추출 가능

```java
        List<Member> members = repository.getAllMember();
        OutputStream os = new FileOutputStream("out.xlsx");
        Excelport.toExcel(os, members.iterator());
```

## 4. Excel 메타 정보 설정

---

### 1. @Excel 애노테이션 추가

- Data class(Dto or Vo Class)를 사용하는 경우, 간단하게 @Excel 애노테이션을 추가

```java
        @Excel(name = "age", order = 2, align = Align.RIGHT)
        int age;
```

- Excel 애노테이션 상세
    - name : 컬럼 타이틀, 디폴트 멤버 변수명
    - order : 컬럼 출력 순서, 낮은 숫자일수록 먼저 출력, 디폴트 999
    - align : 좌우 정렬. 디폴트 None
    - format : 날짜 출력 포맷, 날짜 타입의 객체인경우만 유효

### 2. Excel 추출 메타 정보 추가

- List<Map<String, Object>> 형태의 데이터인 경우, 별도의 메타 정보가 필요
- String Array 타입으로 정의
    - 한 필드의 정보는 아래와 같이 Key=Value 형태로 표현
    - 속성이 여러개인 경우 컴마로 구분하여 표기
    - fieldName 은 Map 에서 key 로 사용될 데이터임. **필수**

```java
        // excel 출력
        String[] template = {
                "fieldName=name,name=이름, order=1",
                "fieldName=age, name=나이, align=right",
                "fieldName=nick_name, name=별명, order=2",
                "fieldName=my_point, name=포인트, align=LEFT",
                "fieldName=joinedAt, dateFormat=YYYYMMDD"
        };
        OutputStream excelOs = new FileOutputStream(new File(tempDir, "out_with_map_string_template.xlsx"));
        Excelport.toExcel(excelOs, list.iterator(), template);
``` 



## 4. Excel 추출 샘플

---

![excelport-result](https://user-images.githubusercontent.com/5078531/141324070-4b3fc604-eb38-4fa9-9ef3-3736f0ecd267.png)


## 5. 기타

---

### 1. 출력값에 대한 서식, 변환등

- getter 메소드가 존재하는 경우, getter 메소드를 호출하여 값을 셋팅한다.
- 따라서, 출력시 출력값을 조절하고 싶은 경우, getter 메소드에 로직 구현.
- getter 메소드가 없는경우, 당연히 필드값 자체를 출력 값으로 셋팅한다.

### 2. 추가 기능

- json export 
    - Excelport.toJson(OutputStream, Iterator) 

- csv export
    - Excelport.toCsv(OutputStream, Iterator)

### 3. mybatis 를 사용하는 경우 

- mybatis 를 사용하면서, OOM(out of memory)를 피하고 싶은경우, Cursor 를 사용해서 fetch 
- 단, Cursor 를 사용하는 경우, 시작과 종료시점까지 Transaction 안에서 사용되어야 한다.

```java
# 일반적인 list 타입의 fetch
List<Member> selectSuperLargeDataSetAsList(Map<String, Object> params);
 
# cursor를 사용한 fetch
Cursor<Member> selectSuperLargeDataSetAsCursor(Map<String, Object> params);
```

### 4. web application 에서 excel download

- web application 에서 excel 다운로드시 HttpServletResponse.getOutputStream 을 이용
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
