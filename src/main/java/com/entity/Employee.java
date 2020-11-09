package com.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.math.BigDecimal;
import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class Employee {
    private String name;
    private String age;
    private Date birthDate;
    private BigDecimal payment;
    private BigDecimal bonus;
}
