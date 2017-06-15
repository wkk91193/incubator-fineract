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
package org.apache.fineract.infrastructure.bulkimport.populator.group;

import org.apache.fineract.infrastructure.bulkimport.populator.AbstractWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.CenterSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ClientSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.PersonnelSheetPopulator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class GroupsWorkbookPopulator extends AbstractWorkbookPopulator {

	private static final int NAME_COL = 0;
	private static final int OFFICE_NAME_COL = 1;
	private static final int STAFF_NAME_COL = 2;
	private static final int CENTER_NAME_COL = 3;
	private static final int EXTERNAL_ID_COL = 4;
	private static final int ACTIVE_COL = 5;
	private static final int ACTIVATION_DATE_COL = 6;
	private static final int MEETING_START_DATE_COL = 7;
	private static final int IS_REPEATING_COL = 8;
	private static final int FREQUENCY_COL = 9;
	private static final int INTERVAL_COL = 10;
	private static final int REPEATS_ON_DAY_COL = 11;
	private static final int STATUS_COL = 12;
	private static final int GROUP_ID_COL = 13;
	private static final int FAILURE_COL = 14;
	private static final int CLIENT_NAMES_STARTING_COL = 15;
	private static final int CLIENT_NAMES_ENDING_COL = 250;
	private static final int LOOKUP_OFFICE_NAME_COL = 251;
	private static final int LOOKUP_OFFICE_OPENING_DATE_COL = 252;
	private static final int LOOKUP_REPEAT_NORMAL_COL = 253;
	private static final int LOOKUP_REPEAT_MONTHLY_COL = 254;
	private static final int LOOKUP_IF_REPEAT_WEEKLY_COL = 255;
	private OfficeSheetPopulator officeSheetPopulator;
	private PersonnelSheetPopulator personnelSheetPopulator;
	private CenterSheetPopulator centerSheetPopulator;
	private ClientSheetPopulator clientSheetPopulator;

	public GroupsWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator,
			PersonnelSheetPopulator personnelSheetPopulator, CenterSheetPopulator centerSheetPopulator,
			ClientSheetPopulator clientSheetPopulator) {
		this.officeSheetPopulator = officeSheetPopulator;
		this.personnelSheetPopulator = personnelSheetPopulator;
		this.centerSheetPopulator = centerSheetPopulator;
		this.clientSheetPopulator = clientSheetPopulator;
	}

	@Override
	public void populate(Workbook workbook) {
		Sheet groupSheet = workbook.createSheet("Groups");
		personnelSheetPopulator.populate(workbook);
		officeSheetPopulator.populate(workbook);
		centerSheetPopulator.populate(workbook);
		clientSheetPopulator.populate(workbook);
		setLayout(groupSheet);
		setLookupTable(groupSheet);

	}

	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
		rowHeader.setHeight((short) 500);
		worksheet.setColumnWidth(NAME_COL, 4000);
		worksheet.setColumnWidth(OFFICE_NAME_COL, 5000);
		worksheet.setColumnWidth(STAFF_NAME_COL, 5000);
		worksheet.setColumnWidth(CENTER_NAME_COL, 5000);
		worksheet.setColumnWidth(EXTERNAL_ID_COL, 2500);
		worksheet.setColumnWidth(ACTIVE_COL, 2000);
		worksheet.setColumnWidth(ACTIVATION_DATE_COL, 3500);
		worksheet.setColumnWidth(MEETING_START_DATE_COL, 3500);
		worksheet.setColumnWidth(IS_REPEATING_COL, 2000);
		worksheet.setColumnWidth(FREQUENCY_COL, 3000);
		worksheet.setColumnWidth(INTERVAL_COL, 2000);
		worksheet.setColumnWidth(REPEATS_ON_DAY_COL, 2500);
		worksheet.setColumnWidth(STATUS_COL, 2000);
		worksheet.setColumnWidth(GROUP_ID_COL, 2000);
		worksheet.setColumnWidth(FAILURE_COL, 2000);
		worksheet.setColumnWidth(CLIENT_NAMES_STARTING_COL, 4000);
		worksheet.setColumnWidth(LOOKUP_OFFICE_NAME_COL, 6000);
		worksheet.setColumnWidth(LOOKUP_OFFICE_OPENING_DATE_COL, 4000);
		worksheet.setColumnWidth(LOOKUP_REPEAT_NORMAL_COL, 3000);
		worksheet.setColumnWidth(LOOKUP_REPEAT_MONTHLY_COL, 3000);
		worksheet.setColumnWidth(LOOKUP_IF_REPEAT_WEEKLY_COL, 3000);

		writeString(NAME_COL, rowHeader, "Group Name*");
		writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
		writeString(STAFF_NAME_COL, rowHeader, "Staff Name*");
		writeString(CENTER_NAME_COL, rowHeader, "Center Name");
		writeString(EXTERNAL_ID_COL, rowHeader, "External ID");
		writeString(ACTIVE_COL, rowHeader, "Active*");
		writeString(ACTIVATION_DATE_COL, rowHeader, "Activation Date*");
		writeString(MEETING_START_DATE_COL, rowHeader, "Meeting Start Date* (On or After)");
		writeString(IS_REPEATING_COL, rowHeader, "Repeat*");
		writeString(FREQUENCY_COL, rowHeader, "Frequency*");
		writeString(INTERVAL_COL, rowHeader, "Interval*");
		writeString(REPEATS_ON_DAY_COL, rowHeader, "Repeats On*");
		writeString(CLIENT_NAMES_STARTING_COL, rowHeader, "Client Names* (Enter in consecutive cells horizontally)");
		writeString(LOOKUP_OFFICE_NAME_COL, rowHeader, "Office Name");
		writeString(LOOKUP_OFFICE_OPENING_DATE_COL, rowHeader, "Opening Date");
		writeString(LOOKUP_REPEAT_NORMAL_COL, rowHeader, "Repeat Normal Range");
		writeString(LOOKUP_REPEAT_MONTHLY_COL, rowHeader, "Repeat Monthly Range");
		writeString(LOOKUP_IF_REPEAT_WEEKLY_COL, rowHeader, "If Repeat Weekly Range");

	}
    private void setLookupTable(Sheet groupSheet) {
    	setOfficeDateLookupTable(groupSheet, officeSheetPopulator.getOffices(), LOOKUP_OFFICE_NAME_COL, LOOKUP_OFFICE_OPENING_DATE_COL);
    	int rowIndex;
    	for(rowIndex = 1; rowIndex <= 11; rowIndex++) {
    		Row row = groupSheet.getRow(rowIndex);
    		if(row == null)
    			row = groupSheet.createRow(rowIndex);
    		writeInt(LOOKUP_REPEAT_MONTHLY_COL, row, rowIndex);
    	}
    	for(rowIndex = 1; rowIndex <= 3; rowIndex++) 
    		writeInt(LOOKUP_REPEAT_NORMAL_COL, groupSheet.getRow(rowIndex), rowIndex);
    	String[] days = new String[]{"Mon","Tue","Wed","Thu","Fri","Sat","Sun"};
    	for(rowIndex = 1; rowIndex <= 7; rowIndex++) 
    		writeString(LOOKUP_IF_REPEAT_WEEKLY_COL, groupSheet.getRow(rowIndex), days[rowIndex-1]);
    }

}
