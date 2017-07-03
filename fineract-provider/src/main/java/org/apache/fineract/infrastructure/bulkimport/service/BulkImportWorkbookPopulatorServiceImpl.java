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

import org.apache.fineract.infrastructure.bulkimport.populator.CenterSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.ClientSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.PersonnelSheetPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.WorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.centers.CentersWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.client.ClientWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.group.GroupsWorkbookPopulator;
import org.apache.fineract.infrastructure.core.exception.GeneralPlatformDomainRuleException;
import org.apache.fineract.infrastructure.core.service.DateUtils;
import org.apache.fineract.infrastructure.security.service.PlatformSecurityContext;
import org.apache.fineract.organisation.office.data.OfficeData;
import org.apache.fineract.organisation.office.service.OfficeReadPlatformService;
import org.apache.fineract.organisation.staff.data.StaffData;
import org.apache.fineract.organisation.staff.service.StaffReadPlatformService;
import org.apache.fineract.portfolio.client.api.ClientApiConstants;
import org.apache.fineract.portfolio.client.data.ClientData;
import org.apache.fineract.portfolio.client.service.ClientReadPlatformService;
import org.apache.fineract.portfolio.group.api.GroupingTypesApiConstants;
import org.apache.fineract.portfolio.group.data.CenterData;
import org.apache.fineract.portfolio.group.service.CenterReadPlatformService;
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
  
  @Autowired
  public BulkImportWorkbookPopulatorServiceImpl(final PlatformSecurityContext context,
      final OfficeReadPlatformService officeReadPlatformService,
      final StaffReadPlatformService staffReadPlatformService,
      final ClientReadPlatformService clientReadPlatformService,
      final CenterReadPlatformService centerReadPlatformService) {
    this.officeReadPlatformService = officeReadPlatformService;
    this.staffReadPlatformService = staffReadPlatformService;
    this.context = context;
    this.clientReadPlatformService=clientReadPlatformService;
    this.centerReadPlatformService=centerReadPlatformService;
    
  }

  @Override
  public Response getClientTemplate(final String entityType, final Long officeId, final Long staffId) {

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

  @SuppressWarnings("unchecked")
  private List<OfficeData> fetchOffices(final Long officeId) {
    List<OfficeData> offices = null;
    if (officeId == null) {
      Boolean includeAllOffices = Boolean.TRUE;
      offices = (List) this.officeReadPlatformService.retrieveAllOffices(includeAllOffices, null);
     // System.out.println("Offices List size : "+offices.size());
    } else {
      offices = new ArrayList<>();
      offices.add(this.officeReadPlatformService.retrieveOffice(officeId));
    }
    return offices;
  }

  @SuppressWarnings("unchecked")
  private List<StaffData> fetchStaff(final Long staffId) {
    List<StaffData> staff = null;
    if (staffId == null){
      staff =
          (List) this.staffReadPlatformService.retrieveAllStaff(null, null, Boolean.FALSE, null);
    //System.out.println("Staff List size : "+staff.size());
    }else {
      staff = new ArrayList<>();
      staff.add(this.staffReadPlatformService.retrieveStaff(staffId));
    }
    return staff;
  }
  
  @Override
  public Response getCentersTemplate(final String entityType, final Long officeId, final Long staffId){
	  
	  WorkbookPopulator populator=null;
	  final Workbook workbook=new HSSFWorkbook();
	  if (entityType.trim().equalsIgnoreCase(GroupingTypesApiConstants.CENTER_RESOURCE_NAME)) {
			populator=populateCenterWorkbook(officeId,staffId);
		} else {
			throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
	          "Unable to find requested resource");
		}
	  	populator.populate(workbook);
	  	return buildResponse(workbook, entityType);
  }
  
private WorkbookPopulator populateCenterWorkbook(Long officeId,Long staffId){
	 this.context.authenticatedUser().validateHasReadPermission("OFFICE");
	 this.context.authenticatedUser().validateHasReadPermission("STAFF");
	 List<OfficeData> offices = fetchOffices(officeId);
	  List<StaffData> staff = fetchStaff(staffId);
	return new CentersWorkbookPopulator(new OfficeSheetPopulator(offices),
	        new PersonnelSheetPopulator(staff, offices));
}

@Override
	public Response getGroupsTemplate(String entityType, Long officeId, Long staffId,Long centerId, Long clientId){
		WorkbookPopulator populator = null;
		final Workbook workbook = new HSSFWorkbook();
		if (entityType.trim().equalsIgnoreCase(GroupingTypesApiConstants.GROUP_RESOURCE_NAME)) {
			populator = populateGroupsWorkbook(officeId, staffId,centerId,clientId);
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
		List<CenterData>centers=null;
		if (centerId==null) {
			centers=(List<CenterData>) this.centerReadPlatformService.retrieveAll(null, null);
		} else {
			centers=new ArrayList<>();
			centers.add(this.centerReadPlatformService.retrieveOne(centerId));
		}
		
		return centers;
	}
	private List<ClientData> fetchClients(Long clientId) {
		List<ClientData> clients=null;
		if (clientId==null) {
			clients=(List<ClientData>) this.clientReadPlatformService.retrieveAllClients();
		} else {
		clients=new ArrayList<>();
		clients.add(this.clientReadPlatformService.retrieveOne(clientId));
		}
		return clients;
	}
}