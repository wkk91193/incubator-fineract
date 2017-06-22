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
package org.apache.fineract.infrastructure.bulkimport.populator;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.fineract.organisation.office.data.OfficeData;
import org.apache.fineract.portfolio.group.data.GroupGeneralData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class GroupSheetPopulator extends AbstractWorkbookPopulator {
	private List<GroupGeneralData> allGroups;
	private List<OfficeData> allOffices;
	// private List<GroupGeneralData> activeGroups;

	private Map<String, ArrayList<String>> officeToGroups;
	private Map<Integer, Integer[]> officeNameToBeginEndIndexesOfGroups;
	private Map<String, Long> groupNameToGroupId;

	private static final int OFFICE_NAME_COL = 0;
	private static final int GROUP_NAME_COL = 1;
	private static final int GROUP_ID_COL = 2;

	public GroupSheetPopulator(final List<GroupGeneralData> groups, final List<OfficeData> offices) {
		this.allGroups = groups;
		this.allOffices = offices;
	}

	@Override
	public void populate(Workbook workbook) {
		Sheet groupSheet = workbook.createSheet("Groups");
		setLayout(groupSheet);
		// filterActiveGroups();
		// System.out.println("Active groups size : " + activeGroups.size());
		setOfficeToGroupsMap();
		setGroupNametoGroupIdMap();
		populateGroupsByOfficeName(groupSheet);
		groupSheet.protectSheet("");
	}

	// private void filterActiveGroups() {
	// activeGroups = new ArrayList<>();
	// groupNameToGroupId = new HashMap<String, Long>();
	// for (GroupGeneralData groupGeneralData : allGroups) {
	// if (groupGeneralData.getActive() != null) {
	// if (groupGeneralData.getActive()) {
	// activeGroups.add(groupGeneralData);
	// groupNameToGroupId.put(groupGeneralData.getName().trim(),
	// groupGeneralData.getId());
	// }
	// }
	//
	// }
	//
	// }
	private void setGroupNametoGroupIdMap() {
		groupNameToGroupId = new HashMap<String, Long>();
		for (GroupGeneralData groupGeneralData : allGroups) {
			groupNameToGroupId.put(groupGeneralData.getName().trim(), groupGeneralData.getId());
		}

	}

	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
		rowHeader.setHeight((short) 500);
		for (int colIndex = 0; colIndex <= 10; colIndex++)
			worksheet.setColumnWidth(colIndex, 6000);
		writeString(OFFICE_NAME_COL, rowHeader, "Office Names");
		writeString(GROUP_NAME_COL, rowHeader, "Group Names");
		writeString(GROUP_ID_COL, rowHeader, "Group ID");
	}

	private void setOfficeToGroupsMap() {
		officeToGroups = new HashMap<String, ArrayList<String>>();
		// for(GroupGeneralData group :activeGroups){
		for (GroupGeneralData group : allGroups) {
			add(group.getOfficeName().trim().replaceAll("[ )(]", "_"), group.getName().trim());
		}
	}

	// Guava Multi-map can reduce this.
	private void add(String key, String value) {
		ArrayList<String> values = officeToGroups.get(key);
		if (values == null) {
			values = new ArrayList<String>();
		}
		values.add(value);
		officeToGroups.put(key, values);
	}

	private void populateGroupsByOfficeName(Sheet groupSheet) {
		int rowIndex = 1, officeIndex = 0, startIndex = 1;
		officeNameToBeginEndIndexesOfGroups = new HashMap<Integer, Integer[]>();
		Row row = groupSheet.createRow(rowIndex);
		for (OfficeData office : allOffices) {
			startIndex = rowIndex + 1;
			writeString(OFFICE_NAME_COL, row, office.name());
			ArrayList<String> groupsList = new ArrayList<String>();

			if (officeToGroups.containsKey(office.name()))
				groupsList = officeToGroups.get(office.name());

			if (!groupsList.isEmpty()) {
				for (String groupName : groupsList) {
					writeString(GROUP_NAME_COL, row, groupName);
					writeLong(GROUP_ID_COL, row, groupNameToGroupId.get(groupName));
					row = groupSheet.createRow(++rowIndex);
				}
				officeNameToBeginEndIndexesOfGroups.put(officeIndex++, new Integer[] { startIndex, rowIndex });
			} else {
				officeNameToBeginEndIndexesOfGroups.put(officeIndex++, new Integer[] { startIndex, rowIndex + 1 });
			}
		}
	}

	public List<GroupGeneralData> getGroups() {
		return allGroups;
		// return activeGroups;
	}

	public Integer getGroupsSize() {
		return allGroups.size();
		// return activeGroups.size();
	}

	public Map<Integer, Integer[]> getOfficeNameToBeginEndIndexesOfGroups() {
		return officeNameToBeginEndIndexesOfGroups;
	}
}
