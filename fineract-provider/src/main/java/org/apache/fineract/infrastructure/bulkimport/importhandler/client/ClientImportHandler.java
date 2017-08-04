/**
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements. See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership. The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package org.apache.fineract.infrastructure.bulkimport.importhandler.client;

import com.google.gson.GsonBuilder;
import org.apache.fineract.commands.domain.CommandWrapper;
import org.apache.fineract.commands.service.CommandWrapperBuilder;
import org.apache.fineract.commands.service.PortfolioCommandSourceWritePlatformService;
import org.apache.fineract.infrastructure.bulkimport.importhandler.AbstractImportHandler;
import org.apache.fineract.infrastructure.bulkimport.importhandler.helper.DateSerializer;
import org.apache.fineract.infrastructure.core.data.CommandProcessingResult;
import org.apache.fineract.portfolio.client.data.ClientData;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.LocalDate;

import java.util.ArrayList;
import java.util.List;

public class ClientImportHandler extends AbstractImportHandler {
	
    private static final int FIRST_NAME_COL = 0;
    private static final int FULL_NAME_COL = 0;
    private static final int LAST_NAME_COL = 1;
    private static final int MIDDLE_NAME_COL = 2;
    private static final int OFFICE_NAME_COL = 3;
    private static final int STAFF_NAME_COL = 4;
    private static final int EXTERNAL_ID_COL = 5;
    private static final int ACTIVATION_DATE_COL = 6;
    private static final int ACTIVE_COL = 7;
    private static final int STATUS_COL = 8;

    private Workbook workbook;
    private List<ClientData> clients;
    private String clientType;

    public ClientImportHandler(Workbook workbook) {
        this.workbook = workbook;
        this.clients=new ArrayList<ClientData>();
    }

    @Override
    public void readExcelFile() {
        Sheet clientSheet=workbook.getSheet("Clients");
        Integer noOfEntries=getNumberOfRows(clientSheet,0);
        clientType=getClientType(clientSheet);
        for (int rowIndex=1;rowIndex<noOfEntries;rowIndex++){
            Row row;
            try {
                row=clientSheet.getRow(rowIndex);
                if (isNotImported(row,STATUS_COL)){
                    clients.add(readClient(row));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
    private String getClientType(Sheet clientSheet) {
        if(readAsString(FIRST_NAME_COL, clientSheet.getRow(0)).equals("First Name*"))
            return "Individual";
        else
            return "Corporate";
    }
    private ClientData readClient(Row row) {
        String officeName = readAsString(OFFICE_NAME_COL, row);
        Long officeId = getIdByName(workbook.getSheet("Offices"), officeName);
        String staffName = readAsString(STAFF_NAME_COL, row);
        Long staffId = getIdByName(workbook.getSheet("Staff"), staffName);
        String externalId = readAsLong(EXTERNAL_ID_COL, row).toString();
        LocalDate activationDate = readAsDate(ACTIVATION_DATE_COL, row);

        Boolean active = readAsBoolean(ACTIVE_COL, row);
        if (clientType.equals("Individual")) {
            String firstName = readAsString(FIRST_NAME_COL, row);
            String lastName = readAsString(LAST_NAME_COL, row);
            String middleName = readAsString(MIDDLE_NAME_COL, row);
            if (firstName == null || firstName.trim().equals("")) {
                throw new IllegalArgumentException("Name is blank");
            }
            return  new ClientData(row.getRowNum(),firstName,lastName,middleName,activationDate,active,externalId,officeId,staffId);
        } else {
            String fullName = readAsString(FULL_NAME_COL, row);
            if (fullName == null || fullName.trim().equals("")) {
                throw new IllegalArgumentException("Name is blank");
            }
            return new ClientData(row.getRowNum(),fullName, activationDate, active, externalId, officeId, staffId);
        }
    }


    @Override
    public void upload(PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        Sheet clientSheet=workbook.getSheet("Clients");
        for (ClientData client: clients) {
            try {
                GsonBuilder gsonBuilder = new GsonBuilder();
                String payload=gsonBuilder.registerTypeAdapter(LocalDate.class, new DateSerializer()).create().toJson(client);
                final CommandWrapper commandRequest = new CommandWrapperBuilder() //
                        .createClient() //
                        .withJson(payload) //
                        .build(); //
                final CommandProcessingResult result = commandsSourceWritePlatformService.logCommandSource(commandRequest);
                Cell statusCell = clientSheet.getRow(client.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException e) {
                String message = parseStatus(e.getMessage());
                Cell statusCell = clientSheet.getRow(client.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));

            }
        }
        clientSheet.setColumnWidth(STATUS_COL, 15000);
        writeString(STATUS_COL, clientSheet.getRow(0), "Status");
    }
}
