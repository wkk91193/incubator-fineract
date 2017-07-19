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
package org.apache.fineract.infrastructure.bulkimport.exporthandler.dto;

import java.util.Locale;


public class Center {

    private final transient Integer rowIndex;
    private final transient String status;
    private final String dateFormat;
    private final Locale locale;
    private final String name;
    private final String officeId;
    private final String staffId;
    private final String externalId;
    private final String active;
    private final String activationDate;

    public Center(String name, String activationDate, String active, String externalId, String officeId, String staffId, Integer rowIndex, String status) {
        this.name = name;
        this.activationDate = activationDate;
        this.active = active;
        this.externalId = externalId;
        this.officeId = officeId;
        this.staffId = staffId;
        this.rowIndex = rowIndex;
        this.status = status;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public Locale getLocale() {
        return locale;
    }

    public String getName() {
        return name;
    }

    public String getOfficeId() {
        return officeId;
    }

    public String getStaffId() {
        return staffId;
    }

    public String getExternalId() {
        return externalId;
    }

    public String isActive() {
        return active;
    }

    public String getActivationDate() {
        return activationDate;
    }

    public String getStatus() {
        return status;
    }

}