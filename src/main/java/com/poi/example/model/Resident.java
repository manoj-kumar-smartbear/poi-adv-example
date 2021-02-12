package com.poi.example.model;

import lombok.Data;

@Data
public class Resident {
    String name;
    String nationalId;
    String email;
    String mobile;
    int age;
    String address;
    boolean isMarried;
}