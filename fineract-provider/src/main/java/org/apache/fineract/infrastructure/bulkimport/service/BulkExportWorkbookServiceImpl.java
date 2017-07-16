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
package org.apache.fineract.infrastructure.bulkimport.service;

import com.sun.jersey.core.header.FormDataContentDisposition;
import org.apache.fineract.commands.service.PortfolioCommandSourceWritePlatformService;
import org.apache.fineract.infrastructure.bulkimport.exporthandler.DataExportHandler;
import org.apache.fineract.infrastructure.bulkimport.exporthandler.client.ClientDataExportHandler;
import org.apache.fineract.infrastructure.core.exception.GeneralPlatformDomainRuleException;
import org.apache.fineract.portfolio.client.api.ClientApiConstants;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.ws.rs.core.Response;
import java.io.IOException;
import java.io.InputStream;


/**
 * Created by K2 on 7/14/2017.
 */
@Service
public class BulkExportWorkbookServiceImpl implements BulkExportWorkbookService {
    private PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService;

    @Autowired
    public BulkExportWorkbookServiceImpl(final PortfolioCommandSourceWritePlatformService commandsSourceWritePlatformService) {
        this.commandsSourceWritePlatformService = commandsSourceWritePlatformService;
    }

    @Override
    public Response postClientTemplate(String entityType , InputStream inputStream,FormDataContentDisposition fileDetail) {
        try {
            Workbook workbook = new HSSFWorkbook(inputStream);
            DataExportHandler dataExportHandler=null;
            if (entityType.trim().equalsIgnoreCase(ClientApiConstants.CLIENT_RESOURCE_NAME)) {
                dataExportHandler = new ClientDataExportHandler(workbook);
            } else {
                throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
                        "Unable to find requested resource");
            }
            dataExportHandler.readExcelFile();
            dataExportHandler.Upload(commandsSourceWritePlatformService);
        }catch (IOException ex){
            throw new GeneralPlatformDomainRuleException("error.msg.io.exception","IO exception occured with "+fileDetail.getFileName()+" "+ex.getMessage());
        }
        return Response.ok(fileDetail.getFileName()+" uploaded successfully").build();
    }

}
