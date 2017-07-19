/**
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements. See the NOTICE multipartFile
 * distributed with this work for additional information
 * regarding copyright ownership. The ASF licenses this multipartFile
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this multipartFile except in compliance
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
package org.apache.fineract.infrastructure.bulkimport.exporthandler.center;

import com.google.gson.Gson;
import org.apache.fineract.commands.domain.CommandWrapper;
import org.apache.fineract.commands.service.CommandWrapperBuilder;
import org.apache.fineract.commands.service.PortfolioCommandSourceWritePlatformService;
import org.apache.fineract.infrastructure.bulkimport.exporthandler.AbstractDataExportHandler;
import org.apache.fineract.infrastructure.bulkimport.exporthandler.dto.Center;
import org.apache.fineract.infrastructure.bulkimport.exporthandler.dto.Meeting;
import org.apache.fineract.infrastructure.bulkimport.exporthandler.dto.WeeklyMeeting;
import org.apache.fineract.infrastructure.core.data.CommandProcessingResult;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;


public class CenterDataExportHandler extends AbstractDataExportHandler {

    private static final int NAME_COL = 0;
    private static final int OFFICE_NAME_COL = 1;
    private static final int STAFF_NAME_COL = 2;
    private static final int EXTERNAL_ID_COL = 3;
    private static final int ACTIVE_COL = 4;
    private static final int ACTIVATION_DATE_COL = 5;
    private static final int MEETING_START_DATE_COL = 6;
    private static final int IS_REPEATING_COL = 7;
    private static final int FREQUENCY_COL = 8;
    private static final int INTERVAL_COL = 9;
    private static final int REPEATS_ON_DAY_COL = 10;
    private static final int STATUS_COL = 11;
    private static final int CENTER_ID_COL = 12;
    private static final int FAILURE_COL = 13;
    private final Workbook workbook;

    private List<Center> centers;
    private List<Meeting> meetings;

    public CenterDataExportHandler(Workbook workbook) {
        this.workbook = workbook;
        this.centers =   new ArrayList<Center>();
        this.meetings = new ArrayList<Meeting>();
    }

    @Override
    public void readExcelFile() {
        Sheet centersSheet = workbook.getSheet("Centers");
        Integer noOfEntries = getNumberOfRows(centersSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = centersSheet.getRow(rowIndex);
                if(isNotImported(row, STATUS_COL)) {
                    centers.add(readCenter(row));
                    meetings.add(readMeeting(row));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private Meeting readMeeting(Row row) {
        String meetingStartDate = readAsDate(MEETING_START_DATE_COL, row);
        String isRepeating = readAsBoolean(IS_REPEATING_COL, row).toString();
        String frequency = readAsString(FREQUENCY_COL, row);
        frequency = getFrequencyId(frequency);
        String interval = readAsInt(INTERVAL_COL, row);
        String repeatsOnDay = readAsString(REPEATS_ON_DAY_COL, row);
        repeatsOnDay = getRepeatsOnDayId(repeatsOnDay);
        if(meetingStartDate.equals(""))
            return null;
        else {
            if(repeatsOnDay.equals(""))
                return new Meeting(meetingStartDate, isRepeating, frequency, interval, row.getRowNum());
            else
                return new WeeklyMeeting(meetingStartDate, isRepeating, frequency, interval, repeatsOnDay, row.getRowNum());
        }
    }

    private String getFrequencyId(String frequency) {
        if(frequency.equalsIgnoreCase("Daily"))
            frequency = "1";
        else if(frequency.equalsIgnoreCase("Weekly"))
            frequency = "2";
        else if(frequency.equalsIgnoreCase("Monthly"))
            frequency = "3";
        else if(frequency.equalsIgnoreCase("Yearly"))
            frequency = "4";
        return frequency;
    }

    private String getRepeatsOnDayId(String repeatsOnDay) {
        if(repeatsOnDay.equalsIgnoreCase("Mon"))
            repeatsOnDay = "1";
        else if(repeatsOnDay.equalsIgnoreCase("Tue"))
            repeatsOnDay = "2";
        else if(repeatsOnDay.equalsIgnoreCase("Wed"))
            repeatsOnDay = "3";
        else if(repeatsOnDay.equalsIgnoreCase("Thu"))
            repeatsOnDay = "4";
        else if(repeatsOnDay.equalsIgnoreCase("Fri"))
            repeatsOnDay = "5";

        else if(repeatsOnDay.equalsIgnoreCase("Sat"))
            repeatsOnDay = "6";
        else if(repeatsOnDay.equalsIgnoreCase("Sun"))
            repeatsOnDay = "7";
        return repeatsOnDay;
    }


    private Center readCenter(Row row) {
        String status = readAsString(STATUS_COL, row);
        String officeName = readAsString(OFFICE_NAME_COL, row);
        String officeId = getIdByName(workbook.getSheet("Offices"), officeName).toString();
        String staffName = readAsString(STAFF_NAME_COL, row);
        String staffId = getIdByName(workbook.getSheet("Staff"), staffName).toString();
        String externalId = readAsLong(EXTERNAL_ID_COL, row);
        String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
        String active = readAsBoolean(ACTIVE_COL, row).toString();
        String centerName = readAsString(NAME_COL, row);
        if(centerName==null||centerName.equals("")) {
            throw new IllegalArgumentException("Name is blank");
        }
        return new Center(centerName, activationDate, active, externalId, officeId, staffId, row.getRowNum(), status);
    }

    @Override
    public void Upload(PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        Sheet centerSheet = workbook.getSheet("Centers");
        int progressLevel = 0;
        String centerId = "";
        for (int i = 0; i < centers.size(); i++) {
            Row row = centerSheet.getRow(centers.get(i).getRowIndex());
            Cell errorReportCell = row.createCell(FAILURE_COL);
            Cell statusCell = row.createCell(STATUS_COL);
            CommandProcessingResult result=null;
            try {
                String status = centers.get(i).getStatus();
                progressLevel = getProgressLevel(status);

                if(progressLevel == 0)
                {
                    result= uploadCenter(i,commandsSourceWritePlatformService);
                    centerId = result.getGroupId().toString();
                    progressLevel = 1;
                } else
                    centerId = readAsInt(CENTER_ID_COL, centerSheet.getRow(centers.get(i).getRowIndex()));

                if(meetings.get(i) != null)
                    progressLevel = uploadCenterMeeting(result, i,commandsSourceWritePlatformService);

                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (RuntimeException e) {
                System.out.println(e);
                String message = parseStatus(e.getMessage());
                String status = "";

                if(progressLevel == 0)
                    status = "Creation";
                else if(progressLevel == 1)
                    status = "Meeting";
                statusCell.setCellValue(status + " failed.");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));

                if(progressLevel>0)
                    row.createCell(CENTER_ID_COL).setCellValue(Integer.parseInt(centerId));

                errorReportCell.setCellValue(message);
            }
        }
        setReportHeaders(centerSheet);
    }

    private int getProgressLevel(String status) {

        if(status.equals("") || status.equals("Creation failed."))
            return 0;
        else if(status.equals("Meeting failed."))
            return 1;
        return 0;
    }
    private CommandProcessingResult uploadCenter(int rowIndex,PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        String payload = new Gson().toJson(centers.get(rowIndex));
        final CommandWrapper commandRequest = new CommandWrapperBuilder() //
                .createCenter() //
                .withJson(payload) //
                .build(); //
        final CommandProcessingResult result = commandsSourceWritePlatformService.logCommandSource(commandRequest);
        return result;
    }

    private void setReportHeaders(Sheet sheet) {
        writeString(STATUS_COL, sheet.getRow(0), "Status");
        writeString(CENTER_ID_COL, sheet.getRow(0), "Center Id");
        writeString(FAILURE_COL, sheet.getRow(0), "Failure Report");
    }

    private Integer uploadCenterMeeting(CommandProcessingResult result, int rowIndex,PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        Meeting meeting = meetings.get(rowIndex);
        meeting.setCenterId(result.getGroupId().toString());
        meeting.setTitle("centers_" + result.getGroupId().toString() + "_CollectionMeeting");
        String payload = new Gson().toJson(meeting);
        CommandWrapper commandWrapper=new CommandWrapper(result.getOfficeId(),result.getGroupId(),result.getClientId(),result.getLoanId(),result.getSavingsId(),null,null,null,null,null,payload,result.getTransactionId(),result.getProductId(),null,null,null);
        final CommandWrapper commandRequest = new CommandWrapperBuilder() //
                .createCalendar(commandWrapper,"CENTER",result.getGroupId()) //
                .withJson(payload) //
                .build(); //
        final CommandProcessingResult meetingresult = commandsSourceWritePlatformService.logCommandSource(commandRequest);
        return 2;
    }
}
