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

/**
 * Created by K2 on 7/15/2017.
 */
public class CorporateClient extends Client {

    private final String fullname;


    public CorporateClient(String fullname, String activationDate, String active, String externalId, String officeId, String staffId, Integer rowIndex ) {
        super(null, null, null, activationDate, active, externalId, officeId, staffId, rowIndex);
        this.fullname = fullname;
    }

    public String getFullName() {
        return this.fullname;
    }

}
