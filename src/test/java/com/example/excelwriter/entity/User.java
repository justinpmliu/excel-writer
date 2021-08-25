package com.example.excelwriter.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;
import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class User {
    private Integer id;
    private String name;
    private String phone;
    private String email;
    private BigDecimal salary;
    private Date birthday;
    private Boolean married;
    private Double holidays;
}
