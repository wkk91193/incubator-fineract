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
package org.apache.fineract.infrastructure.bulkimport.populator.office;

import org.apache.fineract.infrastructure.bulkimport.populator.AbstractWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.fineract.organisation.office.data.OfficeData;
import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.List;

public class OfficeWorkbookPopulator extends AbstractWorkbookPopulator {
    private  List<OfficeData> offices;

    private static final int OFFICE_NAME_COL = 0;
    private static final int PARENT_OFFICE_NAME_COL = 1;
    private static final int OPENED_ON_COL = 2;
    private static final int EXTERNAL_ID_COL = 3;
    private static final int LOOKUP_OFFICE_COL=7;

    public OfficeWorkbookPopulator(List<OfficeData> offices) {
      this.offices=offices;
    }

    @Override
    public void populate(Workbook workbook) {
       // officeSheetPopulator.populate(workbook);
        Sheet officeSheet=workbook.createSheet("Offices");
        setLayout(officeSheet);
        setLookupTable(officeSheet);
        setRules(officeSheet);
    }

    private void setLookupTable(Sheet officeSheet) {
        int rowIndex=1;
        for (OfficeData office:offices) {
            Row row=officeSheet.createRow(rowIndex);
            writeString(LOOKUP_OFFICE_COL,row,office.name());
            rowIndex++;
        }
    }

    private void setLayout(Sheet worksheet){
        Row rowHeader=worksheet.createRow(0);
        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
        worksheet.setColumnWidth(PARENT_OFFICE_NAME_COL,4000);
        worksheet.setColumnWidth(OPENED_ON_COL,3000);
        worksheet.setColumnWidth(EXTERNAL_ID_COL,3000);
        worksheet.setColumnWidth(LOOKUP_OFFICE_COL,4000);

        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
        writeString(PARENT_OFFICE_NAME_COL, rowHeader, "Parent Office*");
        writeString(OPENED_ON_COL, rowHeader, "Opened On Date*");
        writeString(EXTERNAL_ID_COL, rowHeader, "External Id*");
        writeString(LOOKUP_OFFICE_COL, rowHeader, "Lookup Offices");
    }

    private void setRules(Sheet workSheet){
        CellRangeAddressList parentOfficeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PARENT_OFFICE_NAME_COL, PARENT_OFFICE_NAME_COL);
        CellRangeAddressList OpenedOndateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),OPENED_ON_COL,OPENED_ON_COL);

        DataValidationHelper validationHelper=new HSSFDataValidationHelper((HSSFSheet) workSheet);
        setNames(workSheet);

        DataValidationConstraint parentOfficeNameConstraint=validationHelper.createFormulaListConstraint("Office");
        DataValidationConstraint openDateConstraint=validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.LESS_OR_EQUAL,"=TODAY()",null,"dd/mm/yy");

        DataValidation parentOfficeValidation=validationHelper.createValidation(parentOfficeNameConstraint,parentOfficeNameRange);
        DataValidation  openDateValidation=validationHelper.createValidation(openDateConstraint,OpenedOndateRange);

        workSheet.addValidationData(parentOfficeValidation);
        workSheet.addValidationData(openDateValidation);
    }

    private void setNames(Sheet workSheet) {
        Workbook officeWorkbook=workSheet.getWorkbook();
        Name parentOffice=officeWorkbook.createName();
        parentOffice.setNameName("Office");
        parentOffice.setRefersToFormula("Offices!$H$2:$H$"+(offices.size()+1));
    }
}