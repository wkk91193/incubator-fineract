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
package org.apache.fineract.infrastructure.bulkimport.populator.loanrepayment;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.fineract.infrastructure.bulkimport.populator.AbstractWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ClientSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ExtrasSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.fineract.portfolio.loanaccount.data.LoanAccountData;
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

public class LoanRepaymentWorkbookPopulator extends AbstractWorkbookPopulator {
	private OfficeSheetPopulator officeSheetPopulator;
	private ClientSheetPopulator clientSheetPopulator;
	private ExtrasSheetPopulator extrasSheetPopulator;
	private List<LoanAccountData> allloans;
	
	private static final int OFFICE_NAME_COL = 0;
	private static final int CLIENT_NAME_COL = 1;
	private static final int LOAN_ACCOUNT_NO_COL = 2;
	private static final int PRODUCT_COL = 3;
	private static final int PRINCIPAL_COL = 4;
	private static final int AMOUNT_COL = 5;
	private static final int REPAID_ON_DATE_COL = 6;
	private static final int REPAYMENT_TYPE_COL = 7;
	private static final int ACCOUNT_NO_COL = 8;
	private static final int CHECK_NO_COL = 9;
	private static final int ROUTING_CODE_COL = 10;
	private static final int RECEIPT_NO_COL = 11;
	private static final int BANK_NO_COL = 12;
	private static final int LOOKUP_CLIENT_NAME_COL = 14;
	private static final int LOOKUP_ACCOUNT_NO_COL = 15;
	private static final int LOOKUP_PRODUCT_COL = 16;
	private static final int LOOKUP_PRINCIPAL_COL = 17;
	private static final int LOOKUP_LOAN_DISBURSEMENT_DATE_COL = 18;

	public LoanRepaymentWorkbookPopulator(List<LoanAccountData> loans, OfficeSheetPopulator officeSheetPopulator,
			ClientSheetPopulator clientSheetPopulator, ExtrasSheetPopulator extrasSheetPopulator) {
		this.allloans = loans;
		this.officeSheetPopulator = officeSheetPopulator;
		this.clientSheetPopulator = clientSheetPopulator;
		this.extrasSheetPopulator = extrasSheetPopulator;
	}

	@Override
	public void populate(Workbook workbook) {
		Sheet loanRepaymentSheet = workbook.createSheet("LoanRepayment");
		setLayout(loanRepaymentSheet);
		officeSheetPopulator.populate(workbook);
		clientSheetPopulator.populate(workbook);
		extrasSheetPopulator.populate(workbook);
		populateLoansTable(loanRepaymentSheet);
		setClientAndLoanLookupTable(loanRepaymentSheet,allloans,LOOKUP_CLIENT_NAME_COL,LOOKUP_ACCOUNT_NO_COL,LOOKUP_PRODUCT_COL,
				LOOKUP_PRINCIPAL_COL,LOOKUP_LOAN_DISBURSEMENT_DATE_COL);
		setRules(loanRepaymentSheet);
		setDefaults(loanRepaymentSheet);
	}

	private void setDefaults(Sheet worksheet) {
		try {
			for (Integer rowNo = 1; rowNo < 3000; rowNo++) {
				Row row = worksheet.getRow(rowNo);
				if (row == null)
					row = worksheet.createRow(rowNo);
				writeFormula(PRODUCT_COL, row,
						"IF(ISERROR(VLOOKUP($C" + (rowNo + 1) + ",$P$2:$R$" + (allloans.size() + 1)
								+ ",2,FALSE)),\"\",VLOOKUP($C" + (rowNo + 1) + ",$P$2:$R$" + (allloans.size() + 1)
								+ ",2,FALSE))");
				writeFormula(PRINCIPAL_COL, row,
						"IF(ISERROR(VLOOKUP($C" + (rowNo + 1) + ",$P$2:$R$" + (allloans.size() + 1)
								+ ",3,FALSE)),\"\",VLOOKUP($C" + (rowNo + 1) + ",$P$2:$R$" + (allloans.size() + 1)
								+ ",3,FALSE))");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void setRules(Sheet worksheet) {
		CellRangeAddressList officeNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
				OFFICE_NAME_COL, OFFICE_NAME_COL);
		CellRangeAddressList clientNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(),
				CLIENT_NAME_COL, CLIENT_NAME_COL);
		CellRangeAddressList accountNumberRange = new CellRangeAddressList(1,
				SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_ACCOUNT_NO_COL, LOAN_ACCOUNT_NO_COL);
		CellRangeAddressList repaymentTypeRange = new CellRangeAddressList(1,
				SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAYMENT_TYPE_COL, REPAYMENT_TYPE_COL);
		CellRangeAddressList repaymentDateRange = new CellRangeAddressList(1,
				SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAID_ON_DATE_COL, REPAID_ON_DATE_COL);

		DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet) worksheet);

		setNames(worksheet);

		DataValidationConstraint officeNameConstraint = validationHelper.createFormulaListConstraint("Office");
		DataValidationConstraint clientNameConstraint = validationHelper
				.createFormulaListConstraint("INDIRECT(CONCATENATE(\"Client_\",$A1))");
		DataValidationConstraint accountNumberConstraint = validationHelper.createFormulaListConstraint(
				"INDIRECT(CONCATENATE(\"Account_\",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($B1,\" \",\"_\"),\"(\",\"_\"),\")\",\"_\")))");
		DataValidationConstraint paymentTypeConstraint = validationHelper.createFormulaListConstraint("PaymentTypes");
		DataValidationConstraint repaymentDateConstraint = validationHelper.createDateConstraint(
				DataValidationConstraint.OperatorType.BETWEEN,
				"=VLOOKUP($C1,$P$2:$S$" + (allloans.size() + 1) + ",4,FALSE)", "=TODAY()", "dd/mm/yy");

		DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
		DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
		DataValidation accountNumberValidation = validationHelper.createValidation(accountNumberConstraint,
				accountNumberRange);
		DataValidation repaymentTypeValidation = validationHelper.createValidation(paymentTypeConstraint,
				repaymentTypeRange);
		DataValidation repaymentDateValidation = validationHelper.createValidation(repaymentDateConstraint,
				repaymentDateRange);

		worksheet.addValidationData(officeValidation);
		worksheet.addValidationData(clientValidation);
		worksheet.addValidationData(accountNumberValidation);
		worksheet.addValidationData(repaymentTypeValidation);
		worksheet.addValidationData(repaymentDateValidation);

	}

	private void setNames(Sheet worksheet) {
		ArrayList<String> officeNames = new ArrayList<String>(officeSheetPopulator.getOfficeNames());
		Workbook loanRepaymentWorkbook = worksheet.getWorkbook();
		// Office Names
		Name officeGroup = loanRepaymentWorkbook.createName();
		officeGroup.setNameName("Office");
		officeGroup.setRefersToFormula("Offices!$B$2:$B$" + (officeNames.size() + 1));

		// Clients Named after Offices
		for (Integer i = 0; i < officeNames.size(); i++) {
			Integer[] officeNameToBeginEndIndexesOfClients = clientSheetPopulator
					.getOfficeNameToBeginEndIndexesOfClients().get(i);
			Name name = loanRepaymentWorkbook.createName();
			if (officeNameToBeginEndIndexesOfClients != null) {
				name.setNameName("Client_" + officeNames.get(i).trim().replaceAll("[ )(]", "_"));
				name.setRefersToFormula("Clients!$B$" + officeNameToBeginEndIndexesOfClients[0] + ":$B$"
						+ officeNameToBeginEndIndexesOfClients[1]);
			}
		}

		// Counting clients with active loans and starting and end addresses of
		// cells
		HashMap<String, Integer[]> clientNameToBeginEndIndexes = new HashMap<String, Integer[]>();
		ArrayList<String> clientsWithActiveLoans = new ArrayList<String>();
		ArrayList<String> clientIdsWithActiveLoans = new ArrayList<String>();
		int startIndex = 1, endIndex = 1;
		String clientName = "";
		String clientId = "";
		System.out.println("LoanRepaymentWorkbook allloans size : "+allloans.size());
		for (int i = 0; i < allloans.size(); i++) {
			if (!clientName.equals(allloans.get(i).getClientName())) {
				System.out.println(" clientName not equals(allloans.get(i).getClientName()");
				endIndex = i + 1;
				clientNameToBeginEndIndexes.put(clientName, new Integer[] { startIndex, endIndex });
				startIndex = i + 2;
				clientName = allloans.get(i).getClientName();
				clientId = allloans.get(i).getClientId().toString();
				clientsWithActiveLoans.add(clientName);
				clientIdsWithActiveLoans.add(clientId);
			}
			if (i == allloans.size() - 1) {
				endIndex = i + 2;
				clientNameToBeginEndIndexes.put(clientName, new Integer[] { startIndex, endIndex });
			}
		}
		System.out.println("Clients clientsWithActiveLoans size: "+clientsWithActiveLoans.size()+"clientIdsWithActiveLoans "+clientIdsWithActiveLoans.size() );
		// Account Number Named after Clients
		for (int j = 0; j < clientsWithActiveLoans.size(); j++) {
			Name name = loanRepaymentWorkbook.createName();
			System.out.println("clients with loans : "+clientsWithActiveLoans.get(j).replaceAll(" ", "_") + "_"+ clientIdsWithActiveLoans.get(j) + "_");
			name.setNameName("Account_" + clientsWithActiveLoans.get(j).replaceAll(" ", "_") + "_"
					+ clientIdsWithActiveLoans.get(j) + "_");
			name.setRefersToFormula(
					"LoanRepayment!$P$" + clientNameToBeginEndIndexes.get(clientsWithActiveLoans.get(j))[0] + ":$P$"
							+ clientNameToBeginEndIndexes.get(clientsWithActiveLoans.get(j))[1]);
		}

		// Payment Type Name
		Name paymentTypeGroup = loanRepaymentWorkbook.createName();
		paymentTypeGroup.setNameName("PaymentTypes");
		paymentTypeGroup.setRefersToFormula("Extras!$D$2:$D$" + (extrasSheetPopulator.getPaymentTypesSize() + 1));
	}

	private void populateLoansTable(Sheet loanRepaymentSheet) {
		int rowIndex = 1;
		Row row;
		Workbook workbook = loanRepaymentSheet.getWorkbook();
		CellStyle dateCellStyle = workbook.createCellStyle();
		short df = workbook.createDataFormat().getFormat("dd/mm/yy");
		dateCellStyle.setDataFormat(df);
		SimpleDateFormat outputFormat = new SimpleDateFormat("dd/MM/yyyy");
		SimpleDateFormat inputFormat = new SimpleDateFormat("yyyy-MM-dd");
		Date date = null;
		for (LoanAccountData loan : allloans) {
			row = loanRepaymentSheet.createRow(rowIndex++);
			writeString(LOOKUP_CLIENT_NAME_COL, row, loan.getClientName() + "(" + loan.getClientId() + ")");
			writeLong(LOOKUP_ACCOUNT_NO_COL, row, Long.parseLong(loan.getAccountNo()));
			writeString(LOOKUP_PRODUCT_COL, row, loan.getLoanProductName());
			writeDouble(LOOKUP_PRINCIPAL_COL, row, loan.getPrincipal().doubleValue());
			if (loan.getDisbursementDate() != null) {
				try {
					date = inputFormat.parse(loan.getDisbursementDate().toString());
				} catch (ParseException e) {
					e.printStackTrace();
				}
				writeDate(LOOKUP_LOAN_DISBURSEMENT_DATE_COL, row,
						outputFormat.format(date), dateCellStyle);
			}
		}
	}

	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
		rowHeader.setHeight((short) 500);
		worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
		worksheet.setColumnWidth(CLIENT_NAME_COL, 5000);
		worksheet.setColumnWidth(LOAN_ACCOUNT_NO_COL, 3000);
		worksheet.setColumnWidth(PRODUCT_COL, 4000);
		worksheet.setColumnWidth(PRINCIPAL_COL, 4000);
		worksheet.setColumnWidth(AMOUNT_COL, 4000);
		worksheet.setColumnWidth(REPAID_ON_DATE_COL, 3000);
		worksheet.setColumnWidth(REPAYMENT_TYPE_COL, 3000);
		worksheet.setColumnWidth(ACCOUNT_NO_COL, 3000);
		worksheet.setColumnWidth(CHECK_NO_COL, 3000);
		worksheet.setColumnWidth(RECEIPT_NO_COL, 3000);
		worksheet.setColumnWidth(ROUTING_CODE_COL, 3000);
		worksheet.setColumnWidth(BANK_NO_COL, 3000);
		worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 5000);
		worksheet.setColumnWidth(LOOKUP_ACCOUNT_NO_COL, 3000);
		worksheet.setColumnWidth(LOOKUP_PRODUCT_COL, 3000);
		worksheet.setColumnWidth(LOOKUP_PRINCIPAL_COL, 3700);
		worksheet.setColumnWidth(LOOKUP_LOAN_DISBURSEMENT_DATE_COL, 3700);
		writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
		writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
		writeString(LOAN_ACCOUNT_NO_COL, rowHeader, "Loan Account No.*");
		writeString(PRODUCT_COL, rowHeader, "Product Name");
		writeString(PRINCIPAL_COL, rowHeader, "Principal");
		writeString(AMOUNT_COL, rowHeader, "Amount Repaid*");
		writeString(REPAID_ON_DATE_COL, rowHeader, "Date*");
		writeString(REPAYMENT_TYPE_COL, rowHeader, "Type*");
		writeString(ACCOUNT_NO_COL, rowHeader, "Account No");
		writeString(CHECK_NO_COL, rowHeader, "Check No");
		writeString(RECEIPT_NO_COL, rowHeader, "Receipt No");
		writeString(ROUTING_CODE_COL, rowHeader, "Routing Code");
		writeString(BANK_NO_COL, rowHeader, "Bank No");
		writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Lookup Client");
		writeString(LOOKUP_ACCOUNT_NO_COL, rowHeader, "Lookup Account");
		writeString(LOOKUP_PRODUCT_COL, rowHeader, "Lookup Product");
		writeString(LOOKUP_PRINCIPAL_COL, rowHeader, "Lookup Principal");
		writeString(LOOKUP_LOAN_DISBURSEMENT_DATE_COL, rowHeader, "Lookup Loan Disbursement Date");

	}

}