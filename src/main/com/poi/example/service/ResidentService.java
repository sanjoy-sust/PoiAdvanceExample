package main.com.poi.example.service;

import main.com.poi.example.model.Resident;

/**
 * Created by Lenovo on 17/09/2018.
 */
public class ResidentService {

    public Resident getResident() {
        Resident resident = new Resident();
        resident.setName("Flop Coder");
        resident.setAddress("Coding Area");
        resident.setEmail("flop.official@gmail.com");
        resident.setMobile("+8801719149***");
        resident.setAge(38);
        resident.setNationalId("36524585458585545555");
        resident.setIsMarried(true);
        return resident;
    }
}
