package com.example;

import java.io.FileOutputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import io.github.cdimascio.dotenv.Dotenv;

public class Main {
    public static void main(String[] args) {
        

        try {
            Dotenv dotenv = Dotenv.configure().directory("./demo").load();
            String apiToken = dotenv.get("WAZZUP_API_TOKEN");

            HttpClient client = HttpClient.newHttpClient();

            HttpRequest request = HttpRequest.newBuilder()
                    .uri(new URI("https://api.wazzup24.com/v3/templates/whatsapp?limit=250"))
                    .header("Authorization", "Bearer " + apiToken)
                    .GET()
                    .build();

            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() == 200) {
                JSONArray templates = new JSONArray(response.body());

                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Templates");

                Row headerRow = sheet.createRow(0);
                Cell titleHeader = headerRow.createCell(0);
                titleHeader.setCellValue("Title");
                Cell textHeader = headerRow.createCell(1);
                textHeader.setCellValue("Text");

                int rowNum = 1;

                for (int i = 0; i < templates.length(); i++) {
                    JSONObject template = templates.getJSONObject(i);
                    String title = template.getString("title");

                    JSONArray components = template.getJSONArray("components");
                    for (int j = 0; j < components.length(); j++) {
                        JSONObject component = components.getJSONObject(j);
                        if (component.has("text")) {
                            String text = component.getString("text");

                            Row row = sheet.createRow(rowNum++);
                            row.createCell(0).setCellValue(title);
                            row.createCell(1).setCellValue(text);
                        }
                    }
                }

                try (FileOutputStream fileOut = new FileOutputStream("templates.xlsx")) {
                    workbook.write(fileOut);
                }

                workbook.close();
                System.out.println("Excel file created successfully.");

            } else {
                System.out.println("API request failed. HTTP code: " + response.statusCode());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
