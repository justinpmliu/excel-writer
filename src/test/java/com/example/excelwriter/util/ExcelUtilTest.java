package com.example.excelwriter.util;

import com.example.excelwriter.entity.User;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class ExcelUtilTest {

    @Autowired
    private ExcelUtil excelUtil;

    @Test
    void write() throws Exception {
        String format = "yyyy-MM-dd";
        SimpleDateFormat sdf = new SimpleDateFormat(format);

        List<User> users = new ArrayList<>();
        User user = new User(1, "User A", null, "a@mail.com", null,
                sdf.parse("1972-01-02"), true, 17.5);
        users.add(user);

        user = new User(2, "User B", "33333", "b@mail.com", new BigDecimal("20000.99"),
                sdf.parse("1978-06-09"), false, 20.0);
        users.add(user);

        byte[] excelData = excelUtil.listToExcel(users);
        assertNotNull(excelData);

        File file = new File("/home/justin/temp/test.xlsx");
        if (file.exists()) {
            file.delete();
        }
        Files.write(file.toPath(), excelData);
    }

    @Test
    void read() throws Exception {
        File file = new File("/home/justin/temp/test.xlsx");
        byte[] excelData = Files.readAllBytes(file.toPath());
        List<User> users = excelUtil.excelToList(excelData, User.class);
        assertEquals(2, users.size());
        System.out.println(users);
    }
}