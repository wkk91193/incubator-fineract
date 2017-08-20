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
import org.apache.fineract.infrastructure.bulkimport.constants.OfficeConstants;
import org.apache.fineract.infrastructure.bulkimport.constants.TemplatePopulateImportConstants;
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


    public OfficeImportHandler(Workbook workbook) {
        this.offices=new ArrayList<OfficeData>();
        this.workbook=workbook;
    }

    @Override
    public void readExcelFile() {
        Sheet officeSheet=workbook.getSheet(OfficeConstants.OFFICE_WORKBOOK_SHEET_NAME);
        Integer noOfEntries=getNumberOfRows(officeSheet,0);
        for (int rowIndex=1;rowIndex<noOfEntries;rowIndex++){
            Row row;
                row=officeSheet.getRow(rowIndex);
                if (isNotImported(row, OfficeConstants.STATUS_COL)){
                    offices.add(readOffice(row));
                }
        }
    }

    private OfficeData readOffice(Row row) {
        String officeName =readAsString(OfficeConstants.OFFICE_NAME_COL,row);
        Long parentId=readAsLong(OfficeConstants.PARENT_OFFICE_ID_COL,row);
        LocalDate openedDate=readAsDate(OfficeConstants.OPENED_ON_COL,row);
        String externalId=readAsLong(OfficeConstants.EXTERNAL_ID_COL,row).toString();
        return new OfficeData(officeName,parentId,openedDate,externalId,row.getRowNum());
    }

    @Override
    public void Upload(PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        Sheet clientSheet=workbook.getSheet(OfficeConstants.OFFICE_WORKBOOK_SHEET_NAME);
        GsonBuilder gsonBuilder = new GsonBuilder();
        gsonBuilder.registerTypeAdapter(LocalDate.class, new DateSerializer());
        for (OfficeData office: offices) {
            try {
                String payload=gsonBuilder.create().toJson(office);
                final CommandWrapper commandRequest = new CommandWrapperBuilder() //
                        .createOffice() //
                        .withJson(payload) //
                        .build(); //
                final CommandProcessingResult result = commandsSourceWritePlatformService.logCommandSource(commandRequest);
                Cell statusCell = clientSheet.getRow(office.getRowIndex()).createCell(OfficeConstants.STATUS_COL);
                statusCell.setCellValue(TemplatePopulateImportConstants.STATUS_CELL_IMPORTED);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException e) {
                e.printStackTrace();
                String message="";
                if (e.getMessage()!=null)
                 message = parseStatus(e.getMessage());
                Cell statusCell = clientSheet.getRow(office.getRowIndex()).createCell(OfficeConstants.STATUS_COL);
                statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));

            }
        }
        clientSheet.setColumnWidth(OfficeConstants.STATUS_COL, TemplatePopulateImportConstants.SMALL_COL_SIZE);
        writeString(OfficeConstants.STATUS_COL, clientSheet.getRow(0), TemplatePopulateImportConstants.STATUS_COLUMN_HEADER_NAME);
    }
}
