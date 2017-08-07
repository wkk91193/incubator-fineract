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
package org.apache.fineract.infrastructure.bulkimport.importhandler.office;

import com.google.gson.GsonBuilder;
import org.apache.fineract.commands.domain.CommandWrapper;
import org.apache.fineract.commands.service.CommandWrapperBuilder;
import org.apache.fineract.commands.service.PortfolioCommandSourceWritePlatformService;
import org.apache.fineract.infrastructure.bulkimport.importhandler.AbstractImportHandler;
import org.apache.fineract.infrastructure.bulkimport.importhandler.helper.DateSerializer;
import org.apache.fineract.infrastructure.core.data.CommandProcessingResult;
import org.apache.fineract.organisation.office.data.OfficeData;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.LocalDate;

import java.util.ArrayList;
import java.util.List;

public class OfficeImportHandler extends AbstractImportHandler {
    private final List<OfficeData> offices;
    private final Workbook workbook;

    private static final int OFFICE_NAME=0;
    private static final int PARENT_OFFICE=1;
    private static final int PARENT_OFFICE_ID=2;
    private static final int OPENED_ON_DATE=3;
    private static final int  EXTERNAL_ID=4;
    private static final int STATUS_COL=5;

    public OfficeImportHandler(Workbook workbook) {
        this.offices=new ArrayList<OfficeData>();
        this.workbook=workbook;
    }

    @Override
    public void readExcelFile() {
        Sheet officeSheet=workbook.getSheet("Offices");
        Integer noOfEntries=getNumberOfRows(officeSheet,0);
        for (int rowIndex=1;rowIndex<noOfEntries;rowIndex++){
            Row row;
            try {
                row=officeSheet.getRow(rowIndex);
                if (isNotImported(row,STATUS_COL)){
                    offices.add(readOffice(row));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private OfficeData readOffice(Row row) {
        String officeName =readAsString(OFFICE_NAME,row);
        Long parentId=readAsLong(PARENT_OFFICE_ID,row);
        LocalDate openedDate=readAsDate(OPENED_ON_DATE,row);
        String externalId=readAsLong(EXTERNAL_ID,row).toString();
        return new OfficeData(officeName,parentId,openedDate,externalId,row.getRowNum());
    }

    @Override
    public void Upload(PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        Sheet clientSheet=workbook.getSheet("Offices");
        for (OfficeData office: offices) {
            try {
                GsonBuilder gsonBuilder = new GsonBuilder();
                String payload=gsonBuilder.registerTypeAdapter(LocalDate.class, new DateSerializer()).create().toJson(office);
                final CommandWrapper commandRequest = new CommandWrapperBuilder() //
                        .createOffice() //
                        .withJson(payload) //
                        .build(); //
                final CommandProcessingResult result = commandsSourceWritePlatformService.logCommandSource(commandRequest);
                Cell statusCell = clientSheet.getRow(office.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException e) {
                e.printStackTrace();
                String message="";
                if (e.getMessage()!=null)
                 message = parseStatus(e.getMessage());
                Cell statusCell = clientSheet.getRow(office.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));

            }
        }
        clientSheet.setColumnWidth(STATUS_COL, 15000);
        writeString(STATUS_COL, clientSheet.getRow(0), "Status");
    }
}
