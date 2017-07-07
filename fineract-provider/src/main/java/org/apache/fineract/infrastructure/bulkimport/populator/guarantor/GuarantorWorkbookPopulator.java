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
package org.apache.fineract.infrastructure.bulkimport.populator.guarantor;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;

import org.apache.fineract.infrastructure.bulkimport.populator.AbstractWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ClientSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.fineract.portfolio.client.data.ClientData;
import org.apache.fineract.portfolio.loanaccount.data.LoanAccountData;
import org.apache.fineract.portfolio.savings.data.SavingsAccountData;
import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;

public class GuarantorWorkbookPopulator extends AbstractWorkbookPopulator {
	private OfficeSheetPopulator officeSheetPopulator;
	private ClientSheetPopulator clientSheetPopulator;
	private List<LoanAccountData> loans;
	private List<SavingsAccountData> savings;
	
	private static final int OFFICE_NAME_COL = 0;
    private static final int CLIENT_NAME_COL = 1;
    private static final int LOAN_ACCOUNT_NO_COL = 2;
    private static final int GUARANTO_TYPE_COL =3;
    private static final int CLIENT_RELATIONSHIP_TYPE_COL =4;
    private static final int ENTITY_OFFICE_NAME_COL = 5;
    private static final int ENTITY_ID_COL = 6;
    private static final int FIRST_NAME_COL = 7;
    private static final int LAST_NAME_COL = 8;
    private static final int ADDRESS_LINE_1_COL= 9;
    private static final int ADDRESS_LINE_2_COL = 10;
    private static final int CITY_COL = 11;
    private static final int DOB_COL = 12;
    private static final int ZIP_COL = 13;
    private static final int SAVINGS_ID_COL=14;
    private static final int AMOUNT=15;
    private static final int LOOKUP_CLIENT_NAME_COL=81;
    private static final int LOOKUP_ACCOUNT_NO_COL=82;
    private static final int LOOKUP_SAVINGS_CLIENT_NAME_COL=83;
    private static final int LOOKUP_SAVINGS_ACCOUNT_NO_COL=84;

	
	public GuarantorWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator,
			ClientSheetPopulator clientSheetPopulator, List<LoanAccountData> loans, List<SavingsAccountData> savings) {
		this.officeSheetPopulator = officeSheetPopulator;
		this.clientSheetPopulator = clientSheetPopulator;
		this.loans = loans;
		this.savings = savings;
	}

	@Override
	public void populate(Workbook workbook) {
		Sheet addGuarantorSheet = workbook.createSheet("guarantor");
    	setLayout(addGuarantorSheet);
    	officeSheetPopulator.populate(workbook);
    	clientSheetPopulator.populate(workbook);
    	populateLoansTable(addGuarantorSheet);
    	populateSavingsTable(addGuarantorSheet);
    	setRules(addGuarantorSheet);

	}
	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
		 worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
	        worksheet.setColumnWidth(CLIENT_NAME_COL, 5000);
	        worksheet.setColumnWidth(LOAN_ACCOUNT_NO_COL, 3000);
	        worksheet.setColumnWidth(GUARANTO_TYPE_COL, 3000);
	        worksheet.setColumnWidth(CLIENT_RELATIONSHIP_TYPE_COL, 3000);
	        worksheet.setColumnWidth(ENTITY_OFFICE_NAME_COL, 4000);
	        worksheet.setColumnWidth(ENTITY_ID_COL, 3000);
	        worksheet.setColumnWidth(FIRST_NAME_COL, 3000);
	        worksheet.setColumnWidth(LAST_NAME_COL, 3000);
	        worksheet.setColumnWidth(ADDRESS_LINE_1_COL, 3000);
	        worksheet.setColumnWidth(ADDRESS_LINE_2_COL, 3000);
	        worksheet.setColumnWidth(CITY_COL, 3000);
	        worksheet.setColumnWidth(DOB_COL, 3000);
	        worksheet.setColumnWidth(ZIP_COL, 3000);
	        worksheet.setColumnWidth(SAVINGS_ID_COL, 3000);
	        worksheet.setColumnWidth(AMOUNT, 3000);
	        worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 3000);
	        worksheet.setColumnWidth(LOOKUP_ACCOUNT_NO_COL, 3000);
	        worksheet.setColumnWidth(LOOKUP_SAVINGS_CLIENT_NAME_COL, 3000);
	        worksheet.setColumnWidth(LOOKUP_SAVINGS_ACCOUNT_NO_COL, 3000);
	        writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
	        writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
	        writeString(LOAN_ACCOUNT_NO_COL, rowHeader, " Loan Account NO");
	        writeString(GUARANTO_TYPE_COL, rowHeader, "Guranter_type*");
	        writeString(CLIENT_RELATIONSHIP_TYPE_COL, rowHeader, "Client Relationship type*");
	        writeString(ENTITY_OFFICE_NAME_COL, rowHeader, "Guranter office");
	        writeString(ENTITY_ID_COL, rowHeader, "Gurantor client id*");
	        writeString(FIRST_NAME_COL, rowHeader, "First Name*");
	        writeString(LAST_NAME_COL, rowHeader, "Last Name");
	        writeString(ADDRESS_LINE_1_COL, rowHeader, "ADDRESS LINE 1");
	        writeString(ADDRESS_LINE_2_COL, rowHeader, "ADDRESS LINE 2");
	        writeString(CITY_COL, rowHeader, "City");
	        writeString(DOB_COL, rowHeader, "Date of Birth");
	        writeString(ZIP_COL, rowHeader, "Zip*");
	        writeString(SAVINGS_ID_COL, rowHeader, "Savings Account Id");
	        writeString(AMOUNT, rowHeader, "Amount");
	        writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Lookup Client");
	        writeString(LOOKUP_ACCOUNT_NO_COL, rowHeader, "Lookup Loan Account");
	        writeString(LOOKUP_SAVINGS_CLIENT_NAME_COL, rowHeader, "Savings Lookup Client");
	        writeString(LOOKUP_SAVINGS_ACCOUNT_NO_COL, rowHeader, "Savings Lookup Account");
	
	}
    private void populateSavingsTable(Sheet addGuarantorSheet) {
        Workbook workbook = addGuarantorSheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
        int rowIndex = 1;
        Row row;
        Collections.sort(savings, SavingsAccountData.ClientNameComparator);
            for(SavingsAccountData savingsAccount : savings) {
                if(addGuarantorSheet.getRow(rowIndex)==null) {
                    row = addGuarantorSheet.createRow(rowIndex++);
                }
                else {
                 row=addGuarantorSheet.getRow(rowIndex++);
                }
                writeString(LOOKUP_SAVINGS_CLIENT_NAME_COL, row, savingsAccount.getClientName()  + "(" + savingsAccount.getClientId() + ")");
                writeLong(LOOKUP_SAVINGS_ACCOUNT_NO_COL, row, Long.parseLong(savingsAccount.getAccountNo()));
            }

    }
    private void populateLoansTable(Sheet addGuarantorSheet) {
        Workbook workbook = addGuarantorSheet.getWorkbook();
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
        int rowIndex = 1;
        Row row;
        Collections.sort(loans, LoanAccountData.ClientNameComparator);
            for(LoanAccountData loan : loans) {
                if(addGuarantorSheet.getRow(rowIndex)==null){
                    row = addGuarantorSheet.createRow(rowIndex++);
                }
                else{
                    row= addGuarantorSheet.getRow(rowIndex++);
                }
                writeString(LOOKUP_CLIENT_NAME_COL, row, loan.getClientName() + "(" + loan.getClientId() + ")");
                writeLong(LOOKUP_ACCOUNT_NO_COL, row, Long.parseLong(loan.getAccountNo()));
            }
    }

	private void setRules(Sheet worksheet) {

    		CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
        	CellRangeAddressList clientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), CLIENT_NAME_COL, CLIENT_NAME_COL);
        	CellRangeAddressList entityofficeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ENTITY_OFFICE_NAME_COL, ENTITY_OFFICE_NAME_COL);
        	CellRangeAddressList entityclientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ENTITY_ID_COL, ENTITY_ID_COL);
        	CellRangeAddressList accountNumberRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_ACCOUNT_NO_COL, LOAN_ACCOUNT_NO_COL);
        	CellRangeAddressList savingsaccountNumberRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), SAVINGS_ID_COL, SAVINGS_ID_COL);
        	CellRangeAddressList guranterTypeRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GUARANTO_TYPE_COL, GUARANTO_TYPE_COL);
        	
        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
        	
        	setNames(worksheet);
        	
        	DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$A1))");
        	DataValidationConstraint accountNumberConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Account_\",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($B1,\" \",\"_\"),\"(\",\"_\"),\")\",\"_\")))");
        	DataValidationConstraint savingsaccountNumberConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"SavingsAccount_\",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($G1,\" \",\"_\"),\"(\",\"_\"),\")\",\"_\")))");
        	DataValidationConstraint guranterTypeConstraint = validationHelper.createExplicitListConstraint(new String[] {"Internal","External"});
        	DataValidationConstraint entityofficeNameConstraint = validationHelper.createFormulaListConstraint("Office");
        	DataValidationConstraint entityclientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$F1))");
    	
        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
        	DataValidation accountNumberValidation = validationHelper.createValidation(accountNumberConstraint, accountNumberRange);
        	DataValidation savingsaccountNumberValidation = validationHelper.createValidation(savingsaccountNumberConstraint, savingsaccountNumberRange);
        	
        	
        	DataValidation guranterTypeValidation = validationHelper.createValidation(guranterTypeConstraint, guranterTypeRange);
        	DataValidation entityofficeValidation = validationHelper.createValidation(entityofficeNameConstraint, entityofficeNameRange);
        	DataValidation entityclientValidation = validationHelper.createValidation(entityclientNameConstraint, entityclientNameRange);
    	
        	
        	worksheet.addValidationData(officeValidation);
            worksheet.addValidationData(clientValidation);
            worksheet.addValidationData(accountNumberValidation);
            worksheet.addValidationData(guranterTypeValidation);
            worksheet.addValidationData(entityofficeValidation);
            worksheet.addValidationData(entityclientValidation);
            worksheet.addValidationData(savingsaccountNumberValidation);
    	
	}
	private void setNames(Sheet worksheet) {
    	Workbook addGurarantorWorkbook = worksheet.getWorkbook();
    	ArrayList<String> officeNames = new ArrayList<String>(officeSheetPopulator.getOfficeNames());
    	
    	//Office Names
    	Name officeGroup = addGurarantorWorkbook.createName();
    	officeGroup.setNameName("Office");
    	officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));
    	
    	//Clients Named after Offices
    	for(Integer i = 0; i < officeNames.size(); i++) {
    		Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator.getOfficeNameToBeginEndIndexesOfClients().get(i);
    		Name name = addGurarantorWorkbook.createName();
    		if(officeNameToBeginEndIndexesOfClients != null) {
    	       name.setNameName("Client_" + officeNames.get(i).trim().replaceAll("[ )(]", "_"));
    	       name.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$" + officeNameToBeginEndIndexesOfClients[1]);
    		}
    	}
    	
    	//Counting clients with active loans and starting and end addresses of cells
    	HashMap<String, Integer[]> clientNameToBeginEndIndexes = new HashMap<String, Integer[]>();
    	ArrayList<String> clientsWithActiveLoans = new ArrayList<String>();
    	ArrayList<String> clientIdsWithActiveLoans = new ArrayList<String>();
    	int startIndex = 1, endIndex = 1;
    	String clientName = "";
    	String clientId = "";
    	for(int i = 0; i < loans.size(); i++){
    		if(!clientName.equals(loans.get(i).getClientName())) {
    			endIndex = i + 1;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    			startIndex = i + 2;
    			clientName = loans.get(i).getClientName();
    			clientId = loans.get(i).getClientId().toString();
    			clientsWithActiveLoans.add(clientName);
    			clientIdsWithActiveLoans.add(clientId);
    		}
    		if(i == loans.size()-1) {
    			endIndex = i + 2;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    		}
    	}
    	
    	//Account Number Named  after Clients
    	for(int j = 0; j < clientsWithActiveLoans.size(); j++) {
    		Name name = addGurarantorWorkbook.createName();
    		name.setNameName("Account_" + clientsWithActiveLoans.get(j).replaceAll(" ", "_") + "_" + clientIdsWithActiveLoans.get(j) + "_");
    		name.setRefersToFormula("guarantor!$CE$" + clientNameToBeginEndIndexes.get(clientsWithActiveLoans.get(j))[0] + ":$CE$" + clientNameToBeginEndIndexes.get(clientsWithActiveLoans.get(j))[1]);
    	}
    	///savings
    	//Counting clients with active savings and starting and end addresses of cells for naming
    	ArrayList<String> clientsWithActiveSavings = new ArrayList<String>();
    	ArrayList<String> clientIdsWithActiveSavings = new ArrayList<String>();
    	clientName="";
    	clientId="";
    	for(int i = 0; i < savings.size(); i++){
    		if(!clientName.equals(savings.get(i).getClientName())) {
    			endIndex = i + 1;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    			startIndex = i + 2;
    			clientName = savings.get(i).getClientName();
    			clientId = savings.get(i).getClientId().toString();
    			clientsWithActiveSavings.add(clientName);
    			clientIdsWithActiveSavings.add(clientId);
    		}
    		if(i == savings.size()-1) {
    			endIndex = i + 2;
    			clientNameToBeginEndIndexes.put(clientName, new Integer[]{startIndex, endIndex});
    		}
    	}
    	//Account Number Named  after Clients
    	for(int j = 0; j < clientsWithActiveSavings.size(); j++) {
    		Name name = addGurarantorWorkbook.createName();
    		name.setNameName("SavingsAccount_" + clientsWithActiveSavings.get(j).replaceAll(" ", "_") + "_" + clientIdsWithActiveSavings.get(j) + "_");
    		name.setRefersToFormula("guarantor!$CG$" + clientNameToBeginEndIndexes.get(clientsWithActiveSavings.get(j))[0] + ":$CG$" + clientNameToBeginEndIndexes.get(clientsWithActiveSavings.get(j))[1]);
    	}
  	
	}

}