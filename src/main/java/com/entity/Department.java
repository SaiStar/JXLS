package com.entity;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class Department {
    List<Employee> staff = new ArrayList<>();
    private String name;
    private Employee chief;
    private String link;

}
