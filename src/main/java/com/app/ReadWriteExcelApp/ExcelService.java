package com.app.ReadWriteExcelApp;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelService {
    public static void writeToExcelSheet() {

        String[] row_heading = {"First Name","Last Name","Address","Email"};

        List<User> users = createUserData();

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet( "User Details ");
        Row headerRow = spreadsheet.createRow(0);

        // Creating header
        for (int i = 0; i < row_heading.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(row_heading[i]);
        }
        // Creating data rows for each user
        for(int i = 0; i < users.size(); i++) {
            Row dataRow = spreadsheet.createRow(i + 1);
            dataRow.createCell(0).setCellValue(users.get(i).getFirstName());
            dataRow.createCell(1).setCellValue(users.get(i).getLastName());
            dataRow.createCell(2).setCellValue(users.get(i).getAddress());
            dataRow.createCell(3).setCellValue(users.get(i).getEmail());
        }

        //Write the workbook in file system
        FileOutputStream out;
        try {
            out = new FileOutputStream( new File("F:/writeToExcelSheet/user.xlsx"));

            workbook.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


        System.out.println("Write to excel sheet done  successfully...........");

    }


    public static  List<User> createUserData() {

        List<User> users = new ArrayList<User>();
        users.add(new User("Lipsa", "Patra", "BBSR","abc@gmail.com" ));
        users.add(new User("Ravish", "Sharma", "Banglore","ravi@gmail.com"));
        users.add(new User("Julia", "Robert",  "Amsterdam","robert@gmail.com"));
        users.add(new User("Meghna", "Morkel", "London","megha@gmail.com"));
        users.add(new User("Morish", "Harison",  "USA","mharison@yahoo.co.in"));
        return users;
    }



    public List<User> ReadDataFromExcel(String path) {

        List<User> userList = new ArrayList<User>();

        try {
            XSSFWorkbook workbook = new XSSFWorkbook(path);

            // Retrieving the number of sheets in the Workbook
            System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
            System.out.println("Retrieving Sheets using for-each loop");
            for(Sheet sheet: workbook) {


                int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
                System.out.println("rowCount........ "  +  rowCount);
                for (int i=1; i<=rowCount; i++) {
                    Row row = sheet.getRow(i);
                    System.out.println("no of rows.... "  +  row.getRowNum() );
                    String firstName = row.getCell(0).getStringCellValue();
                    String lastName = row.getCell(1).getStringCellValue();
                    String email = row.getCell(2).getStringCellValue();
                    String address = row.getCell(3).getStringCellValue();

                    System.out.println("firstName........ "  + firstName);
                    System.out.println("lastName........ "  + lastName);
                    System.out.println("email........ "  + email);
                    System.out.println("address........ "  + address);

                    User user = new User();
                    user.setFirstName(firstName);
                    user.setLastName(lastName);
                    user.setAddress(address);
                    user.setEmail(email);

                    userList.add(user);
                }

            }
        }catch (IOException e) {
            e.printStackTrace();
        }
        return userList;



    }
}
