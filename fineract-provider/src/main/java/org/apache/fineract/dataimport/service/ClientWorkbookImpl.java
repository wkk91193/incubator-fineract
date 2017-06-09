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
package org.apache.fineract.dataimport.service;

import java.util.List;

import org.apache.fineract.dataimport.dboperations.ClientDbOperations;
import org.apache.fineract.dataimport.dto.client.Client;
import org.apache.fineract.dataimport.handler.WorkbookUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class ClientWorkbookImpl implements ClientWorkbook {
	private static final int FIRST_NAME_COL = 0;
    private static final int LAST_NAME_COL = 1;
    private static final int MIDDLE_NAME_COL = 2;
    private static final int FULL_NAME_COL = 0;
    private static final int OFFICE_ID = 3;
    private static final int STAFF_ID = 4;
    private static final int EXTERNAL_ID_COL = 5;
    private static final int ACTIVATION_DATE_COL = 6;
	private String clientType;
	private ClientDbOperations clientDbOperations;
	
	@Autowired
	ClientWorkbookImpl(ClientDbOperations clientDbOperations){
		this.clientDbOperations=clientDbOperations;
	}


	@Override
	public Workbook getTemplate(String clientType) {
		HSSFWorkbook workbook=new HSSFWorkbook();
		Sheet sheet=workbook.createSheet("Clients");
		this.clientType=clientType;
		setLayout(sheet);
		populate(sheet);
		return workbook;
	}

	private void populate(Sheet sheet){
		List<Client> clientList=clientDbOperations.getClientData();	
		int rownum=1;
		if(clientType.equals("individual")){			
			for (Client client : clientList) {
				Row row=sheet.createRow(rownum);			
				WorkbookUtils.writeString(FIRST_NAME_COL, row,client.getFirstname());
				WorkbookUtils.writeString(LAST_NAME_COL, row,client.getLastname());
	            WorkbookUtils.writeString(MIDDLE_NAME_COL, row,client.getMiddlename());
	            WorkbookUtils.writeString(OFFICE_ID, row, client.getOfficeId());
	            WorkbookUtils.writeString(STAFF_ID, row, client.getStaffId());
	            WorkbookUtils.writeString(EXTERNAL_ID_COL, row, client.getExternalId());
	            WorkbookUtils.writeString(ACTIVATION_DATE_COL, row,client.getActivationDate());
	            rownum++;
			}
		}else{
			for (Client client : clientList) {
				Row row=sheet.createRow(rownum);			
				WorkbookUtils.writeString(FULL_NAME_COL, row,client.getFullname());
				WorkbookUtils.writeString(LAST_NAME_COL, row,client.getLastname());
	            WorkbookUtils.writeString(MIDDLE_NAME_COL, row,client.getMiddlename());
	            WorkbookUtils.writeString(OFFICE_ID, row, client.getOfficeId());
	            WorkbookUtils.writeString(STAFF_ID, row, client.getStaffId());
	            WorkbookUtils.writeString(EXTERNAL_ID_COL, row, client.getExternalId());
	            WorkbookUtils.writeString(ACTIVATION_DATE_COL, row,client.getActivationDate());
	            rownum++;
			}
		}		
	}
	@Override
	public void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
    	if(clientType.equals("individual")) {
    	    worksheet.setColumnWidth(FIRST_NAME_COL, 6000);
            worksheet.setColumnWidth(LAST_NAME_COL, 6000);
            worksheet.setColumnWidth(MIDDLE_NAME_COL, 6000);
            WorkbookUtils.writeString(FIRST_NAME_COL, rowHeader, "First Name*");
            WorkbookUtils.writeString(LAST_NAME_COL, rowHeader, "Last Name*");
            WorkbookUtils.writeString(MIDDLE_NAME_COL, rowHeader, "Middle Name");
    	} else {
    		worksheet.setColumnWidth(FULL_NAME_COL, 10000);
    		worksheet.setColumnWidth(LAST_NAME_COL, 0);
    		worksheet.setColumnWidth(MIDDLE_NAME_COL, 0);
    		WorkbookUtils.writeString(FULL_NAME_COL, rowHeader, "Full/Business Name*");
    		WorkbookUtils.writeString(LAST_NAME_COL, rowHeader, "Last Name*");
            WorkbookUtils.writeString(MIDDLE_NAME_COL, rowHeader, "Middle Name");
    	}
        worksheet.setColumnWidth(OFFICE_ID, 5000);
        worksheet.setColumnWidth(STAFF_ID, 5000);
        worksheet.setColumnWidth(EXTERNAL_ID_COL, 3500);
        worksheet.setColumnWidth(ACTIVATION_DATE_COL, 4000);
        WorkbookUtils.writeString(OFFICE_ID, rowHeader, "Office ID*");
        WorkbookUtils.writeString(STAFF_ID, rowHeader, "Staff ID*");
        WorkbookUtils.writeString(EXTERNAL_ID_COL, rowHeader, "External ID");
        WorkbookUtils.writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date*");
		
	}

}
