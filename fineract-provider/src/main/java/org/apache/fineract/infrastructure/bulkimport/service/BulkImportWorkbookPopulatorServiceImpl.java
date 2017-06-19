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
package org.apache.fineract.infrastructure.bulkimport.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;
import org.apache.fineract.accounting.glaccount.api.GLAccountsApiConstants;
import org.apache.fineract.accounting.glaccount.data.GLAccountData;
import org.apache.fineract.accounting.glaccount.service.GLAccountReadPlatformService;
import org.apache.fineract.accounting.journalentry.api.JournalEntriesApiConstants;

import org.apache.fineract.infrastructure.bulkimport.populator.CenterSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ClientSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ExtrasSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.GroupSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.LoanProductSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.PersonnelSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.WorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.centers.CentersWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.client.ClientWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.glaccount.GLAccountWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.group.GroupsWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.loan.LoanWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.loanrepayment.LoanRepaymentWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.office.OfficeWorkbookPopulator;
import org.apache.fineract.infrastructure.core.exception.GeneralPlatformDomainRuleException;
import org.apache.fineract.infrastructure.core.service.DateUtils;
import org.apache.fineract.infrastructure.security.service.PlatformSecurityContext;
import org.apache.fineract.organisation.monetary.data.CurrencyData;
import org.apache.fineract.organisation.monetary.service.CurrencyReadPlatformService;
import org.apache.fineract.organisation.office.api.OfficeApiConstants;
import org.apache.fineract.organisation.office.data.OfficeData;
import org.apache.fineract.organisation.office.service.OfficeReadPlatformService;
import org.apache.fineract.organisation.staff.data.StaffData;
import org.apache.fineract.organisation.staff.service.StaffReadPlatformService;
import org.apache.fineract.portfolio.client.api.ClientApiConstants;
import org.apache.fineract.portfolio.client.data.ClientData;
import org.apache.fineract.portfolio.client.service.ClientReadPlatformService;
import org.apache.fineract.portfolio.fund.data.FundData;
import org.apache.fineract.portfolio.fund.service.FundReadPlatformService;
import org.apache.fineract.portfolio.group.api.GroupingTypesApiConstants;
import org.apache.fineract.portfolio.group.data.CenterData;
import org.apache.fineract.portfolio.group.data.GroupGeneralData;
import org.apache.fineract.portfolio.group.service.CenterReadPlatformService;
import org.apache.fineract.portfolio.group.service.GroupReadPlatformService;
import org.apache.fineract.portfolio.loanaccount.api.LoanApiConstants;
import org.apache.fineract.portfolio.loanaccount.data.LoanAccountData;
import org.apache.fineract.portfolio.loanaccount.service.LoanReadPlatformService;
import org.apache.fineract.portfolio.loanproduct.data.LoanProductData;
import org.apache.fineract.portfolio.loanproduct.service.LoanProductReadPlatformService;
import org.apache.fineract.portfolio.loanrepayment.api.LoanRepaymentApiConstants;
import org.apache.fineract.portfolio.paymenttype.data.PaymentTypeData;
import org.apache.fineract.portfolio.paymenttype.service.PaymentTypeReadPlatformService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class BulkImportWorkbookPopulatorServiceImpl implements BulkImportWorkbookPopulatorService {

	private final PlatformSecurityContext context;
	private final OfficeReadPlatformService officeReadPlatformService;
	private final StaffReadPlatformService staffReadPlatformService;
	private final ClientReadPlatformService clientReadPlatformService;
	private final CenterReadPlatformService centerReadPlatformService;
	private final GroupReadPlatformService groupReadPlatformService;
	private final FundReadPlatformService fundReadPlatformService;
	private final PaymentTypeReadPlatformService paymentTypeReadPlatformService;
	private final LoanProductReadPlatformService loanProductReadPlatformService;
	private final CurrencyReadPlatformService currencyReadPlatformService;
	private final LoanReadPlatformService loanReadPlatformService;

	@Autowired
	public BulkImportWorkbookPopulatorServiceImpl(final PlatformSecurityContext context,
			final OfficeReadPlatformService officeReadPlatformService,
			final StaffReadPlatformService staffReadPlatformService,
			final ClientReadPlatformService clientReadPlatformService,
			final CenterReadPlatformService centerReadPlatformService,
			final GroupReadPlatformService groupReadPlatformService,
			final FundReadPlatformService fundReadPlatformService,
			final PaymentTypeReadPlatformService paymentTypeReadPlatformService,
			final LoanProductReadPlatformService loanProductReadPlatformService,
			final CurrencyReadPlatformService currencyReadPlatformService,
			final LoanReadPlatformService loanReadPlatformService) {
		this.officeReadPlatformService = officeReadPlatformService;
		this.staffReadPlatformService = staffReadPlatformService;
		this.context = context;
		this.clientReadPlatformService = clientReadPlatformService;
		this.centerReadPlatformService = centerReadPlatformService;
		this.groupReadPlatformService = groupReadPlatformService;
		this.fundReadPlatformService = fundReadPlatformService;
		this.paymentTypeReadPlatformService = paymentTypeReadPlatformService;
		this.loanProductReadPlatformService = loanProductReadPlatformService;
		this.currencyReadPlatformService = currencyReadPlatformService;
		this.loanReadPlatformService=loanReadPlatformService;
	}

	@Override
	public Response getClientsTemplate(final String entityType, final Long officeId, final Long staffId) {

		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(ClientApiConstants.CLIENT_RESOURCE_NAME)) {
			populator = populateClientWorkbook(officeId, staffId);
		} else
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
					"Unable to find requested resource");
		populator.populate(workbook);
		return buildResponse(workbook, entityType);
	}

	@Override
	public Response getOfficesTemplate(final String entityType, final Long officeId) {
		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(OfficeApiConstants.OFFICE_RESOURCE_NAME)) {
			populator = populateOfficeWorkbook(officeId);
		} else {
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
					"Unable to find requested resource");
		}
		populator.populate(workbook);
		return buildResponse(workbook, entityType);
	}

	@Override
	public Response getCentersTemplate(final String entityType, final Long officeId, final Long staffId) {

		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(GroupingTypesApiConstants.CENTER_RESOURCE_NAME)) {
			populator = populateCenterWorkbook(officeId, staffId);
		} else {
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
					"Unable to find requested resource");
		}
		populator.populate(workbook);
		return buildResponse(workbook, entityType);
	}

	private WorkbookPopulator populateCenterWorkbook(Long officeId, Long staffId) {
		this.context.authenticatedUser().validateHasReadPermission("OFFICE");
		this.context.authenticatedUser().validateHasReadPermission("STAFF");
		List<OfficeData> offices = fetchOffices(officeId);
		List<StaffData> staff = fetchStaff(staffId);
		return new CentersWorkbookPopulator(new OfficeSheetPopulator(offices),
				new PersonnelSheetPopulator(staff, offices));
	}

	private WorkbookPopulator populateOfficeWorkbook(Long officeId) {
		this.context.authenticatedUser().validateHasReadPermission("OFFICE");
		List<OfficeData> offices = fetchOffices(officeId);
		return new OfficeWorkbookPopulator(new OfficeSheetPopulator(offices));
	}

	private WorkbookPopulator populateClientWorkbook(final Long officeId, final Long staffId) {
		this.context.authenticatedUser().validateHasReadPermission("OFFICE");
		this.context.authenticatedUser().validateHasReadPermission("STAFF");
		List<OfficeData> offices = fetchOffices(officeId);
		List<StaffData> staff = fetchStaff(staffId);

		return new ClientWorkbookPopulator(new OfficeSheetPopulator(offices),
				new PersonnelSheetPopulator(staff, offices));
	}

	private Response buildResponse(final Workbook workbook, final String entity) {
		String filename = entity + DateUtils.getLocalDateOfTenant().toString() + ".xls";
		final ByteArrayOutputStream baos = new ByteArrayOutputStream();
		try {
			workbook.write(baos);
		} catch (IOException e) {
			e.printStackTrace();
		}

		final ResponseBuilder response = Response.ok(baos.toByteArray());
		response.header("Content-Disposition", "attachment; filename=\"" + filename + "\"");
		response.header("Content-Type", "application/vnd.ms-excel");
		return response.build();
	}

	private List<OfficeData> fetchOffices(final Long officeId) {
		List<OfficeData> offices = null;
		if (officeId == null) {
			Boolean includeAllOffices = Boolean.TRUE;
			offices = (List<OfficeData>) this.officeReadPlatformService.retrieveAllOffices(includeAllOffices, null);
		} else {
			offices = new ArrayList<>();
			offices.add(this.officeReadPlatformService.retrieveOffice(officeId));
		}
		return offices;
	}

	private List<StaffData> fetchStaff(final Long staffId) {
		List<StaffData> staff = null;
		if (staffId == null)
			staff = (List<StaffData>) this.staffReadPlatformService.retrieveAllStaff(null, null, Boolean.FALSE, null);
		else {
			staff = new ArrayList<>();
			staff.add(this.staffReadPlatformService.retrieveStaff(staffId));
		}
		return staff;
	}

	@Override
	public Response getGroupsTemplate(String entityType, Long officeId, Long staffId, Long centerId, Long clientId) {
		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(GroupingTypesApiConstants.GROUP_RESOURCE_NAME)) {
			populator = populateGroupsWorkbook(officeId, staffId, centerId, clientId);
		} else {
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
					"Unable to find requested resource");
		}
		populator.populate(workbook);
		return buildResponse(workbook, entityType);
	}

	private WorkbookPopulator populateGroupsWorkbook(Long officeId, Long staffId, Long centerId, Long clientId) {
		this.context.authenticatedUser().validateHasReadPermission("OFFICE");
		this.context.authenticatedUser().validateHasReadPermission("STAFF");
		this.context.authenticatedUser().validateHasReadPermission("CENTER");
		this.context.authenticatedUser().validateHasReadPermission("CLIENT");
		List<OfficeData> offices = fetchOffices(officeId);
		List<StaffData> staff = fetchStaff(staffId);
		List<CenterData> centers = fetchCenters(centerId);
		List<ClientData> clients = fetchClients(clientId);
		return new GroupsWorkbookPopulator(new OfficeSheetPopulator(offices),
				new PersonnelSheetPopulator(staff, offices), new CenterSheetPopulator(centers, offices),
				new ClientSheetPopulator(clients, offices));
	}

	private List<CenterData> fetchCenters(Long centerId) {
		List<CenterData> centers = null;
		if (centerId == null) {
			centers = (List<CenterData>) this.centerReadPlatformService.retrieveAll(null, null);
		} else {
			centers = new ArrayList<>();
			centers.add(this.centerReadPlatformService.retrieveOne(centerId));
		}

		return centers;
	}

	private List<ClientData> fetchClients(Long clientId) {
		List<ClientData> clients = null;
		if (clientId == null) {
			clients = (List<ClientData>) this.clientReadPlatformService.retrieveAllClients();
		} else {
			clients = new ArrayList<>();
			clients.add(this.clientReadPlatformService.retrieveOne(clientId));
		}
		return clients;
	}

	@Override
	public Response getLoanTemplate(String entityType, Long officeId, Long staffId, Long clientId, Long groupId,
			Long productId, Long fundId, Long paymentTypeId, String code) {
		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(LoanApiConstants.LOAN_RESOURCE_NAME)) {
			populator = populateLoanWorkbook(officeId, staffId, clientId, groupId, productId, fundId, paymentTypeId,
					code);
		} else {
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
					"Unable to find requested resource");
		}
		populator.populate(workbook);
		return buildResponse(workbook, entityType);
	}

	private WorkbookPopulator populateLoanWorkbook(Long officeId, Long staffId, Long clientId, Long groupId,
			Long productId, Long fundId, Long paymentTypeId, String code) {
		this.context.authenticatedUser().validateHasReadPermission("OFFICE");
		this.context.authenticatedUser().validateHasReadPermission("STAFF");
		this.context.authenticatedUser().validateHasReadPermission("GROUP");
		this.context.authenticatedUser().validateHasReadPermission("CLIENT");
		this.context.authenticatedUser().validateHasReadPermission("LOANPRODUCT");
		this.context.authenticatedUser().validateHasReadPermission("FUNDS");
		this.context.authenticatedUser().validateHasReadPermission("PAYMENTTYPE");
		this.context.authenticatedUser().validateHasReadPermission("CURRENCY");
		List<OfficeData> offices = fetchOffices(officeId);
		List<StaffData> staff = fetchStaff(staffId);
		List<ClientData> clients = fetchClients(clientId);
		List<GroupGeneralData> groups = fetchGroups(groupId);
		List<LoanProductData> loanproducts = fetchLoanProducts(productId);
		List<FundData> funds = fetchFunds(fundId);
		List<PaymentTypeData> paymentTypes = fetchPaymentTypes(paymentTypeId);
		List<CurrencyData> currencies = fetchCurrencies(code);
		return new LoanWorkbookPopulator(new OfficeSheetPopulator(offices), new ClientSheetPopulator(clients, offices),
				new GroupSheetPopulator(groups, offices), new PersonnelSheetPopulator(staff, offices),
				new LoanProductSheetPopulator(loanproducts), new ExtrasSheetPopulator(funds, paymentTypes, currencies));
	}

	private List<CurrencyData> fetchCurrencies(String code) {
		List<CurrencyData> currencies = null;
		if (code == null) {
			currencies = (List<CurrencyData>) this.currencyReadPlatformService.retrieveAllPlatformCurrencies();
		} else {
			currencies = new ArrayList<>();
			currencies.add(this.currencyReadPlatformService.retrieveCurrency(code));
		}
		return currencies;
	}

	private List<PaymentTypeData> fetchPaymentTypes(Long paymentTypeId) {
		List<PaymentTypeData> paymentTypeData = null;
		if (paymentTypeId == null) {
			paymentTypeData = (List<PaymentTypeData>) this.paymentTypeReadPlatformService.retrieveAllPaymentTypes();
		} else {
			paymentTypeData = new ArrayList<>();
			paymentTypeData.add(this.paymentTypeReadPlatformService.retrieveOne(paymentTypeId));
		}
		return paymentTypeData;
	}

	private List<FundData> fetchFunds(Long fundId) {
		List<FundData> funds = null;
		if (fundId == null) {
			funds = (List<FundData>) this.fundReadPlatformService.retrieveAllFunds();
		} else {
			funds = new ArrayList<>();
			funds.add(this.fundReadPlatformService.retrieveFund(fundId));
		}
		return funds;
	}

	private List<LoanProductData> fetchLoanProducts(Long productId) {
		List<LoanProductData> loanproducts = null;
		if (productId == null) {
			loanproducts = (List<LoanProductData>) this.loanProductReadPlatformService.retrieveAllLoanProducts();
		} else {
			loanproducts = new ArrayList<>();
			loanproducts.add(this.loanProductReadPlatformService.retrieveLoanProduct(productId));
		}
		return loanproducts;
	}

	private List<GroupGeneralData> fetchGroups(Long groupId) {
		List<GroupGeneralData> groups = null;
		if (groupId == null) {
			groups = (List<GroupGeneralData>) this.groupReadPlatformService.retrieveAll(null, null);
		} else {
			groups = new ArrayList<>();
			groups.add(this.groupReadPlatformService.retrieveOne(groupId));
		}

		return groups;
	}

	@Override
	public Response getLoanRepaymentTemplate(String entityType, Long officeId, Long clientId, Long fundId,
			Long paymentTypeId, String code) {
		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(LoanRepaymentApiConstants.LOAN_REPAYMENT_RESOURCE_NAME)) {
			populator = populateLoanRepaymentWorkbook(officeId, clientId, fundId, paymentTypeId, code);
		} else {
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
					"Unable to find requested resource");
		}
		populator.populate(workbook);
		return buildResponse(workbook, entityType);
	}

	private WorkbookPopulator populateLoanRepaymentWorkbook(Long officeId, Long clientId, Long fundId,
			Long paymentTypeId, String code) {
		this.context.authenticatedUser().validateHasReadPermission("OFFICE");
		this.context.authenticatedUser().validateHasReadPermission("CLIENT");
		this.context.authenticatedUser().validateHasReadPermission("FUNDS");
		this.context.authenticatedUser().validateHasReadPermission("PAYMENTTYPE");
		this.context.authenticatedUser().validateHasReadPermission("CURRENCY");
		List<OfficeData> offices = fetchOffices(officeId);
		List<ClientData> clients = fetchClients(clientId);
		List<FundData> funds = fetchFunds(fundId);
		List<PaymentTypeData> paymentTypes = fetchPaymentTypes(paymentTypeId);
		List<CurrencyData> currencies = fetchCurrencies(code);
		List<LoanAccountData> loans = fetchLoanAccounts();
		return new LoanRepaymentWorkbookPopulator(loans, new OfficeSheetPopulator(offices),
				new ClientSheetPopulator(clients, offices), new ExtrasSheetPopulator(funds, paymentTypes, currencies));
	}

	private List<LoanAccountData> fetchLoanAccounts() {
		List<LoanAccountData> loanaccounts = (List<LoanAccountData>) loanReadPlatformService.retrieveAllLoans();
		return loanaccounts;
	}

}
