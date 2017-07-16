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
package org.apache.fineract.infrastructure.bulkimport.exporthandler.dto;

import java.util.Locale;

/**
 * Created by K2 on 7/15/2017.
 */
public class Client {

    private final transient Integer rowIndex;

    private final String dateFormat;

    private final Locale locale;

    private final String officeId;

    private final String staffId;

    private final String firstname;

    private final String middlename;

    private final String lastname;

    private final String externalId;

    private final String active;

    private final String activationDate;


    public Client(String firstname, String lastname, String middlename, String activationDate, String active, String externalId, String officeId, String staffId, Integer rowIndex ) {
        this.firstname = firstname;
        this.lastname = lastname;
        this.middlename = middlename;
        this.activationDate = activationDate;
        this.active = active;
        this.externalId = externalId;
        this.officeId = officeId;
        this.staffId = staffId;
        this.rowIndex = rowIndex;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
    }

    public String getFirstName() {
        return this.firstname;
    }

    public String getLastName() {
        return this.lastname;
    }

    public String getMiddleName() {
        return this.middlename;
    }

    public String getActivationDate() {
        return this.activationDate;
    }

    public String isActive() {
        return this.active;
    }

    public String getExternalId() {
        return this.externalId;
    }

    public String getOfficeId() {
        return this.officeId;
    }

    public String getStaffId() {
        return this.staffId;
    }

    public Locale getLocale() {
        return locale;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }
}