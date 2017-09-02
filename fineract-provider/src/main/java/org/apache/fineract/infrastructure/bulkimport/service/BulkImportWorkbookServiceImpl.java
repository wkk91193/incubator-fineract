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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLConnection;

import org.apache.fineract.infrastructure.bulkimport.data.BulkImportEvent;
import org.apache.fineract.infrastructure.bulkimport.data.GlobalEntityType;
import org.apache.fineract.infrastructure.bulkimport.data.ImportFormatType;
import org.apache.fineract.infrastructure.bulkimport.domain.ImportDocument;
import org.apache.fineract.infrastructure.bulkimport.domain.ImportDocumentRepository;
import org.apache.fineract.infrastructure.bulkimport.importhandler.ImportHandlerUtils;
import org.apache.fineract.infrastructure.core.exception.GeneralPlatformDomainRuleException;
import org.apache.fineract.infrastructure.core.service.DateUtils;
import org.apache.fineract.infrastructure.core.service.ThreadLocalContextUtil;
import org.apache.fineract.infrastructure.documentmanagement.domain.Document;
import org.apache.fineract.infrastructure.documentmanagement.domain.DocumentRepository;
import org.apache.fineract.infrastructure.documentmanagement.service.DocumentWritePlatformService;
import org.apache.fineract.infrastructure.documentmanagement.service.DocumentWritePlatformServiceJpaRepositoryImpl.DOCUMENT_MANAGEMENT_ENTITY;
import org.apache.fineract.infrastructure.security.service.PlatformSecurityContext;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.tika.Tika;
import org.apache.tika.io.TikaInputStream;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Service;

import com.google.common.io.Files;
import com.sun.jersey.core.header.FormDataContentDisposition;

@Service
public class BulkImportWorkbookServiceImpl implements BulkImportWorkbookService {
	
    private final ApplicationContext applicationContext;
    private final PlatformSecurityContext securityContext;
    private final DocumentWritePlatformService documentWritePlatformService;
    private final DocumentRepository documentRepository;
    private final ImportDocumentRepository importDocumentRepository;

    @Autowired
    public BulkImportWorkbookServiceImpl(final ApplicationContext applicationContext,
    		final PlatformSecurityContext securityContext,
    		final DocumentWritePlatformService documentWritePlatformService,
    		final DocumentRepository documentRepository,
    		final ImportDocumentRepository importDocumentRepository) {
        this.applicationContext = applicationContext;
        this.securityContext = securityContext;
        this.documentWritePlatformService = documentWritePlatformService;
        this.documentRepository = documentRepository;
        this.importDocumentRepository = importDocumentRepository;
    }

    @Override
    public Long importWorkbook(final String entity, final InputStream inputStream,
    		final FormDataContentDisposition fileDetail, final String locale, final String dateFormat) {
        if (entity != null && inputStream != null && fileDetail != null) {
            try {
                final ByteArrayOutputStream baos = new ByteArrayOutputStream();
                IOUtils.copy(inputStream, baos);
                final byte[] bytes = baos.toByteArray();
                InputStream clonedInputStream = new ByteArrayInputStream(bytes);
                InputStream clonedInputStreamWorkbook = new ByteArrayInputStream(bytes);
                final Tika tika = new Tika();
                final TikaInputStream tikaInputStream = TikaInputStream.get(clonedInputStream);
                final String fileType = tika.detect(tikaInputStream);
                final String fileExtension = Files.getFileExtension(fileDetail.getFileName()).toLowerCase();
                if (!fileType.contains("msoffice")) {
                    throw new GeneralPlatformDomainRuleException("error.msg.invalid.file.extension",
                            "Uploaded file extension is not recognized.");
                }

                Workbook workbook = new HSSFWorkbook(clonedInputStreamWorkbook);
                final GlobalEntityType entityType;
                int primaryColumn;
                if (entity.trim().equalsIgnoreCase(GlobalEntityType.OFFICES.toString())) {
                		entityType = GlobalEntityType.OFFICES;
                		primaryColumn = 0;
                } else {
                    throw new GeneralPlatformDomainRuleException("error.msg.unable.to.find.resource",
                            "Unable to find requested resource");
                }
                return publishEvent(primaryColumn, fileDetail, clonedInputStreamWorkbook, entityType,
                		workbook, locale, dateFormat);
                
                
            } catch (IOException ex) {
                throw new GeneralPlatformDomainRuleException("error.msg.io.exception", "IO exception occured with " 
                		+ fileDetail.getFileName() + " " + ex.getMessage());
            }
        } else {
            throw new GeneralPlatformDomainRuleException("error.msg.entityType.null",
                    "Given entityType null or file not found");
        }
    }
    
    private Long publishEvent(final Integer primaryColumn,
    		final FormDataContentDisposition fileDetail, final InputStream clonedInputStreamWorkbook,
    		final GlobalEntityType entityType, final Workbook workbook,
    		final String locale, final String dateFormat) {
    	
    		final String fileName = fileDetail.getFileName();
        final Long documentId = this.documentWritePlatformService.createInternalDocument(
        		DOCUMENT_MANAGEMENT_ENTITY.IMPORT.name(),
        		this.securityContext.authenticatedUser().getId(), null, clonedInputStreamWorkbook,
        		URLConnection.guessContentTypeFromName(fileName), fileName, null, fileName);
        final Document document = this.documentRepository.findOne(documentId);
        
        final ImportDocument importDocument = ImportDocument.instance(document,
        		DateUtils.getLocalDateOfTenant(), entityType.getValue(),
        		this.securityContext.authenticatedUser(),
        		ImportHandlerUtils.getNumberOfRows(workbook.getSheetAt(0),
        				primaryColumn));
        this.importDocumentRepository.saveAndFlush(importDocument);
        BulkImportEvent event = BulkImportEvent.instance(ThreadLocalContextUtil.getTenant()
        		.getTenantIdentifier(), workbook, importDocument.getId(), locale, dateFormat);
        applicationContext.publishEvent(event);
        return importDocument.getId();
    }
}