package com.hdfcbank.service;

import com.hdfcbank.Constant;
import com.hdfcbank.helper.ExcelHelper;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.util.*;

import static com.hdfcbank.helper.ExcelHelper.partyPrimaryInformation;

@Service
public class ExcelService {
//    @Autowired
//    TutorialRepository repository;


    public Map<String,Object> readExcel(MultipartFile file) {
        Map<String,Object> response=new HashMap<>();
        try {
            List<Map<String,Object>> requestList = createJson(file.getInputStream());
            List successList=new ArrayList();
            List errorList=new ArrayList();
//            for (Map<String,Object> objectMap: requestList) {
//                try {
//                    RestTemplate restTemplate = new RestTemplate();
//                    URI uri = new URI("http://apigw-hdfcpoc.centralindia.cloudapp.azure.com/party/party");
//                    HttpHeaders headers = new HttpHeaders();
//                    headers.set("x-requesting-user", "aftab");
//                    headers.set("content-type", "application/json");
//                    HttpEntity<Map> httpEntity = new HttpEntity<>(objectMap, headers);
//                    ResponseEntity<String> result = restTemplate.postForEntity(uri, httpEntity, String.class);
//                    successList.add(result);
//                }
//                catch(Exception e){
//                    errorList.add(e.getMessage()+ " +"+objectMap.get("partyId"));
//                }
//            }
            response.put("request",requestList);
            response.put("Success",successList);
            response.put("Error",errorList);

            // repository.saveAll(tutorials);
        } catch (Exception e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
        return response;
    }

    public static List createJson(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet("party");
            Iterator<Row> rows = sheet.iterator();

            List list = new ArrayList();
            Map<String,Integer> header=new HashMap<>();
            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();


                // skip header
                if (rowNumber == 0) {
                    int rowLength=currentRow.getLastCellNum();
                    for(int i=0; i<rowLength;i++){
                       header.put(currentRow.getCell(i).getStringCellValue(),i);
                    }
                   rowNumber++;
                    continue;
                }
                Map<String,Object> request=new HashMap<>();
                Map<String,Object> headerObject=new HashMap<>();
                Map<String,Object> body=new HashMap<>();
                if (currentRow.getCell(0).getCellType() == CellType.STRING)
                    body.put("guid",String.valueOf(currentRow.getCell(header.get(Constant.UCIC)).getStringCellValue()));
                else
                    body.put("guid",String.valueOf(Long.valueOf((long) currentRow.getCell(header.get(Constant.UCIC)).getNumericCellValue())));
                Map<String,Object> partyPrimaryInformation=new HashMap<>();
                partyPrimaryInformation(currentRow, partyPrimaryInformation,header);

                request.put("partyprimaryInformation",partyPrimaryInformation);
                //request.put("body",body);
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






//    private Mono<ResponseEntity<?>> party(Map<String,Object> objectMap){
//        return createParty(objectMap)
//                .map(responseEntity -> {
//                    if (responseEntity.getStatusCode().is2xxSuccessful()) {
//                        return ResponseEntity.ok(responseEntity.getBody());
//                    } else {
//                        return ResponseEntity.status(responseEntity.getStatusCode())
//                                .body("Failed to create employee");
//                    }
//                })
//                .onErrorResume(exception -> {
//                    return Mono.just(ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
//                            .body("Internal Server Error: " + exception.getMessage()));
//                });
//    }
//    public Mono<Map> createParty(Map<String,Object> request) {
//
//        return client.post()
//                .uri("/party/party")
//                .bodyValue(request)
//                .retrieve()
//                .onStatus(HttpStatus::is4xxClientError, clientResponse -> {
//                    System.out.println(clientResponse);
//                    Map map=new HashMap();
//                    return (Mono<? extends Throwable>) map;
//                })
//                .onStatus(HttpStatus::is5xxServerError, clientResponse -> {
//                    System.out.println(clientResponse);
//                    Map map=new HashMap();
//                    return (Mono<? extends Throwable>) map;
//                })
//                .toEntity(Map.class).flatMap(responseEntity -> Mono.justOrEmpty(responseEntity.getBody()));
//
//    }


//    public ByteArrayInputStream load() {
//        List<Tutorial> tutorials = repository.findAll();
//
//        ByteArrayInputStream in = ExcelHelper.tutorialsToExcel(tutorials);
//        return in;
//    }

//    public List<Tutorial> getAllTutorials() {
//        return repository.findAll();
//    }
}
