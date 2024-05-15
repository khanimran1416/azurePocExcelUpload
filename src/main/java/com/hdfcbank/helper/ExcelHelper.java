package com.hdfcbank.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import com.hdfcbank.Constant;
import com.hdfcbank.model.Tutorial;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;

public class ExcelHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    public static String TYPE1 ="application/vnd.ms-excel";
    static String[] HEADERs = { "Id", "Title", "Description", "Published" };
    static String SHEET = "Tutorials";

    public static boolean hasExcelFormat(MultipartFile file) {

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }

    public static void addHeader(Map<String,Object> header, String key,String value){

        header.put(key,value);
    }

    public static void phoneAddress(Row currentRow, Map<String, Object> phoneAddress,Map<String,Integer> headers){
        phoneAddress.put("extensionFields",null);
        phoneAddress.put("partyType","INDIVIDUAL");
    }
    private void mobilePhoneNumberDetails(Row currentRow, Map<String, Object> phoneAddress,Map<String,Integer> headers){
        Map<String,Object> primaryPhoneAddress=new HashMap<>();
    }
    public static void partyPrimaryInformation(Row currentRow, Map<String, Object> partyPrimaryInformation,Map<String,Integer> headers) {

        partyPrimaryInformation.put("partyType","INDIVIDUAL");

        partyName(currentRow, partyPrimaryInformation,headers);
        familyName(currentRow,partyPrimaryInformation,headers);
        if (currentRow.getCell(headers.get(Constant.GENDER)).getCellType() == CellType.STRING) {
            partyPrimaryInformation.put("gender",String.valueOf(currentRow.getCell(headers.get(Constant.GENDER)).getStringCellValue()));
        }
        if (currentRow.getCell(headers.get(Constant.MARITAL_STATUS)).getCellType() == CellType.STRING) {
            partyPrimaryInformation.put("maritalStatus",String.valueOf(currentRow.getCell(headers.get(Constant.MARITAL_STATUS)).getStringCellValue()));
        }
        if (currentRow.getCell(headers.get(Constant.RESIDENT_TYPE)).getCellType() == CellType.STRING) {
            partyPrimaryInformation.put("residentialStatus",String.valueOf(currentRow.getCell(headers.get(Constant.RESIDENT_TYPE)).getStringCellValue()));
        }

        Map<String,Object> birthInformation=new HashMap<>();
        if (currentRow.getCell(headers.get(Constant.DATE_OF_BIRTH)).getCellType() == CellType.STRING) {
            birthInformation.put("dateOfBirth",String.valueOf(currentRow.getCell(headers.get(Constant.DATE_OF_BIRTH)).getStringCellValue()));
        }
        if (currentRow.getCell(headers.get(Constant.CITY_OF_BIRTH)).getCellType() == CellType.STRING) {
            birthInformation.put("cityOfBirth",String.valueOf(currentRow.getCell(headers.get(Constant.CITY_OF_BIRTH)).getStringCellValue()));
        }
        if (currentRow.getCell(headers.get(Constant.COUNTRY_OF_BIRTH)).getCellType() == CellType.STRING) {
            birthInformation.put("countryOfBirth",String.valueOf(currentRow.getCell(headers.get(Constant.COUNTRY_OF_BIRTH)).getStringCellValue()));
        }
        partyPrimaryInformation.put("birthInformation",birthInformation);
        if (currentRow.getCell(headers.get(Constant.NATIONALITY)).getCellType() == CellType.STRING) {
            partyPrimaryInformation.put("nationality",String.valueOf(currentRow.getCell(headers.get(Constant.NATIONALITY)).getStringCellValue()));
        }
        if (currentRow.getCell(headers.get(Constant.CUSTOMER_SEGMENT)).getCellType() == CellType.STRING) {
            partyPrimaryInformation.put("customerSegment",String.valueOf(currentRow.getCell(headers.get(Constant.CUSTOMER_SEGMENT)).getStringCellValue()));
        }
    }

    private static void familyName(Row currentRow, Map<String, Object> partyPrimaryInformation,Map<String,Integer> headers) {
        Map<String,Object> familyName=new HashMap<>();
        Map<String,Object> fatherName=new HashMap<>();
        Map<String,Object> motherName=new HashMap<>();
        Map<String,Object> spouseName=new HashMap<>();
        if (currentRow.getCell(headers.get(Constant.FATHER_NAME)).getCellType() == CellType.STRING) {
            fatherName.put("fullName", String.valueOf(currentRow.getCell(headers.get(Constant.FATHER_NAME)).getStringCellValue()));
            familyName.put("fatherName",fatherName);
        }
        if (currentRow.getCell(headers.get(Constant.MOTHER_NAME)).getCellType() == CellType.STRING) {
            motherName.put("fullName", String.valueOf(currentRow.getCell(headers.get(Constant.MOTHER_NAME)).getStringCellValue()));
            familyName.put("motherName",motherName);
        }
        if (currentRow.getCell(headers.get(Constant.SPOUSE)).getCellType() == CellType.STRING) {
            spouseName.put("fullName", String.valueOf(currentRow.getCell(headers.get(Constant.SPOUSE)).getStringCellValue()));
            familyName.put("spouseName", spouseName);
        }
        partyPrimaryInformation.put("familyName",familyName);
    }

    private static void partyName(Row currentRow, Map<String, Object> partyPrimaryInformation,Map<String,Integer> headers) {
        Map<String,Object> partyName=new HashMap<>();
        Map<String,Object> name=new HashMap<>();
        name.put("prefix",String.valueOf(currentRow.getCell(headers.get(Constant.SALUTATION)).getStringCellValue()));

        if (currentRow.getCell(headers.get(Constant.CUSTOMER_NAME)).getCellType() == CellType.STRING) {
            String customerName=currentRow.getCell(headers.get(Constant.CUSTOMER_NAME)).getStringCellValue();
            name.put("fullName", customerName);
            String[] str=customerName.split(" ");
            name.put("firstName",str[0]);
            if(str.length<=2){
                name.put("lastName",str[1]);
            }
            else
                name.put("lastName",str[str.length-1]);
        }
        partyName.put("name",name);
        if (currentRow.getCell(headers.get(Constant.CUSTOMER_SHORT_NAME)).getCellType() == CellType.STRING)
        partyName.put("shortName",currentRow.getCell(headers.get(Constant.CUSTOMER_SHORT_NAME)).getStringCellValue());
        partyPrimaryInformation.put("partyName",partyName);
    }
    public static ByteArrayInputStream tutorialsToExcel(List<Tutorial> tutorials) {

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
            Sheet sheet = workbook.createSheet(SHEET);

            // Header
            Row headerRow = sheet.createRow(0);

            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }

            int rowIdx = 1;
            for (Tutorial tutorial : tutorials) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(tutorial.getId());
                row.createCell(1).setCellValue(tutorial.getTitle());
                row.createCell(2).setCellValue(tutorial.getDescription());
                row.createCell(3).setCellValue(tutorial.isPublished());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
        }
    }

    public static List excelToTutorials(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet("Party");
            Iterator<Row> rows = sheet.iterator();

            List list = new ArrayList();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();


                // skip header
                if (rowNumber == 0 || rowNumber == 1) {
                    rowNumber++;
                    continue;
                }

                Map<String,Object> jsonContent=new HashMap<>();

                Map<String,Object> request=new HashMap<>();
                if (currentRow.getCell(0).getCellType() == CellType.STRING)
                request.put("partyId",String.valueOf(currentRow.getCell(0).getStringCellValue()));
                else
                    request.put("partyId",String.valueOf(Long.valueOf((long) currentRow.getCell(0).getNumericCellValue())));
                if (currentRow.getCell(1).getCellType() == CellType.STRING)
                request.put("guid",String.valueOf(currentRow.getCell(1).getStringCellValue()));
                else
                    request.put("guid",String.valueOf(Long.valueOf((long) currentRow.getCell(1).getNumericCellValue())));
                jsonContent.put("partyType",String.valueOf(currentRow.getCell(2).getStringCellValue()));
                if(currentRow.getCell(3) !=null && currentRow.getCell(3).getStringCellValue() !="")
                jsonContent.put("fullName",String.valueOf(currentRow.getCell(3).getStringCellValue()));

                if(currentRow.getCell(4) !=null ){
                        Map<String,Object> phoneNumberDetails=new HashMap<>();
                        if (currentRow.getCell(4).getCellType() == CellType.STRING)
                        phoneNumberDetails.put("phoneNumber",String.valueOf(currentRow.getCell(4).getStringCellValue()));
                        else
                            phoneNumberDetails.put("phoneNumber",String.valueOf(Long.valueOf((long) currentRow.getCell(4).getNumericCellValue())));
                        phoneNumberDetails.put("countryCode","+"+Integer.valueOf((int) currentRow.getCell(5).getNumericCellValue()));
                    jsonContent.put("phoneNumberDetails",phoneNumberDetails);
                }
                List alternatePhoneNumber=new ArrayList();
                if (currentRow.getCell(7) !=null){

                    Map<String,Object> alternatePhoneNumberDetails=new HashMap<>();
                    alternatePhoneNumberDetails.put("mobileType",String.valueOf(currentRow.getCell(6).getStringCellValue()));
                    if (currentRow.getCell(7).getCellType() == CellType.STRING)
                        alternatePhoneNumberDetails.put("phoneNumber",String.valueOf(currentRow.getCell(7).getStringCellValue()));
                        else
                    alternatePhoneNumberDetails.put("phoneNumber",String.valueOf(Long.valueOf((long) currentRow.getCell(7).getNumericCellValue())));
                    alternatePhoneNumberDetails.put("countryCode","+"+Integer.valueOf((int) currentRow.getCell(8).getNumericCellValue()));
                    alternatePhoneNumber.add(alternatePhoneNumberDetails);
                    jsonContent.put("alternatePhoneNumberDetails",alternatePhoneNumber);
                }
                if (currentRow.getCell(9) !=null){
                    jsonContent.put("email",String.valueOf(currentRow.getCell(9).getStringCellValue()));
                }

                if (currentRow.getCell(11) !=null){
                    Map<String,Object> alternateEmailDetails = new HashMap<>();
                    alternateEmailDetails.put("emailType",String.valueOf(currentRow.getCell(10).getStringCellValue()));
                    alternateEmailDetails.put("emailId",String.valueOf(currentRow.getCell(11).getStringCellValue()));
                    List alternateEmailDetailsList=new ArrayList();
                    alternateEmailDetailsList.add(alternateEmailDetails);
                    jsonContent.put("alternateEmailDetails",alternateEmailDetailsList);
                }
                if (currentRow.getCell(12) !=null){
                    jsonContent.put("dateOfBirth",String.valueOf(currentRow.getCell(12).getStringCellValue()));
                }
                if (currentRow.getCell(14) !=null){
                    Map<String, Object> externalReferences=new HashMap<>();
                    externalReferences.put("externalSystem",String.valueOf(currentRow.getCell(13).getStringCellValue()));
                    externalReferences.put("externalId",String.valueOf(currentRow.getCell(14).getStringCellValue()));
                    List externalReferencesList=new ArrayList();
                    externalReferencesList.add(externalReferences);
                    jsonContent.put("externalReferences",externalReferencesList);
                }
                List identifications=new ArrayList();
                if (currentRow.getCell(15) !=null && currentRow.getCell(15).getStringCellValue() !="") {
                    Map aadhaarIdentifications=new HashMap();
                    aadhaarIdentifications.put("type", "AADHAAR");
                    if(currentRow.getCell(16) !=null && currentRow.getCell(16).getStringCellValue() !="")
                    aadhaarIdentifications.put("value", String.valueOf(currentRow.getCell(16).getStringCellValue()));
                    identifications.add(aadhaarIdentifications);
                }
                if (currentRow.getCell(17) !=null && currentRow.getCell(17).getStringCellValue() !="") {
                    Map panIdentifications=new HashMap();
                    panIdentifications.put("type", "INDIVIDUAL_PAN");
                    if (currentRow.getCell(18) !=null && currentRow.getCell(18).getStringCellValue() !="")
                    panIdentifications.put("value", String.valueOf(currentRow.getCell(18).getStringCellValue()));
                    identifications.add(panIdentifications);
                }
                if (currentRow.getCell(19) !=null && currentRow.getCell(19).getStringCellValue() !="") {
                    Map nregaIdentifications=new HashMap();
                    nregaIdentifications.put("type", "NREGA_CARD");
                    if (currentRow.getCell(20) !=null && currentRow.getCell(20).getStringCellValue() !="")
                    nregaIdentifications.put("value", String.valueOf(currentRow.getCell(20).getStringCellValue()));
                    identifications.add(nregaIdentifications);
                }
                if (currentRow.getCell(21) !=null && currentRow.getCell(21).getStringCellValue() !="") {
                    Map voterIdIdentifications=new HashMap();
                    voterIdIdentifications.put("type", "VOTER_ID");
                    if(currentRow.getCell(22) !=null && currentRow.getCell(22).getStringCellValue() !="")
                    voterIdIdentifications.put("value", String.valueOf(currentRow.getCell(22).getStringCellValue()));
                    identifications.add(voterIdIdentifications);
                }
                if (currentRow.getCell(23) !=null && currentRow.getCell(23).getStringCellValue() !="") {
                    Map poiIdentifications=new HashMap();
                    poiIdentifications.put("type", "PASSPORT");
                    if (currentRow.getCell(24) !=null && currentRow.getCell(24).getStringCellValue() !="")
                    poiIdentifications.put("value", String.valueOf(currentRow.getCell(24).getStringCellValue()));
                    identifications.add(poiIdentifications);
                }
//                if (currentRow.getCell(25) !=null && currentRow.getCell(25).getStringCellValue() !="") {
//                    Map frroIdentifications=new HashMap();
//                    frroIdentifications.put("type", "FRRO");
//                    if(currentRow.getCell(26) !=null && currentRow.getCell(26).getStringCellValue() !="")
//                    frroIdentifications.put("value", String.valueOf(currentRow.getCell(26).getStringCellValue()));
//                    identifications.add(frroIdentifications);
//                }

                jsonContent.put("identifications",identifications);
                request.put("jsonContent",jsonContent);
                list.add(request);
            }

            workbook.close();

            return list;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
        catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }
}
