package main;

import main.com.poi.example.model.Resident;
import main.com.poi.example.service.ExcelFileManagerService;
import main.com.poi.example.service.ExcelFileManagerServiceImpl;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * Created by Lenovo on 24I/10/2018.
 */
public class App {
    public static void main(String[] args) {
        ExcelFileManagerService service = new ExcelFileManagerServiceImpl();
        Scanner scanner = new Scanner(System.in);
        System.out.println("For Import press I\n" +
                "For Export press X\n" +
                "For Validate press V");
        String operationType = scanner.next();

        try {
            if (operationType.equals("I")) {
                System.out.println("Importing...");
                System.out.println("Enter how many resident you want to import :");
                int numberOfResident = scanner.nextInt();
                List<Resident> residentList = new ArrayList<>();
                for (int i = 0; i < numberOfResident; i++) {
                    Resident resident = new Resident();
                    System.out.println("Name : ");
                    resident.setName(scanner.next());
                    System.out.println("Address : ");
                    resident.setAddress(scanner.next());
                    System.out.println("Mobile : ");
                    resident.setMobile(scanner.next());
                    System.out.println("Email : ");
                    resident.setEmail(scanner.next());
                    System.out.println("Age : ");
                    resident.setAge(scanner.nextInt());
                    System.out.println("NationalId : ");
                    resident.setNationalId(scanner.next());
                    residentList.add(resident);
                }
                service.importExcel("Imported_Resident.xlsx", residentList);
            } else if (operationType.equals("X")) {
                service.exportExcel("Resident_File.xlsx");
            } else if (operationType.equals("V")) {
                try {
                    service.exportExcel("Resident_File.xlsx");
                    if (service.validateExcelFile("Resident_File.xlsx")) {
                        System.out.println("File is valid");
                    }
                } catch (OpenXML4JException e) {
                    e.printStackTrace();
                } catch (ParserConfigurationException e) {
                    e.printStackTrace();
                } catch (SAXException e) {
                    e.printStackTrace();
                }
            }

        } catch (IOException e) {
            System.out.printf("Not Imported/Exported %s", e.getMessage());
        }
    }
}
