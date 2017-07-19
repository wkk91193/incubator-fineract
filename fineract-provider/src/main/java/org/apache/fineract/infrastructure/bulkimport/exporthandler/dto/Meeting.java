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


public class Meeting {

    private final transient Integer rowIndex;
    private transient String groupId;
    private transient String centerId;
    private final String dateFormat;
    private final Locale locale;
    private final String description;
    private final String typeId;
    private String title;
    private final String startDate;
    private final String repeating;
    private final String frequency;
    private final String interval;

    public Meeting(String startDate, String repeating, String frequency, String interval, Integer rowIndex ) {
        this.startDate = startDate;
        this.repeating = repeating;
        this.frequency = frequency;
        this.interval = interval;
        this.rowIndex = rowIndex;
        this.dateFormat = "dd MMMM yyyy";
        this.locale = Locale.ENGLISH;
        this.description = "";
        this.typeId = "1";
    }

    public void setGroupId(String groupId) {
        this.groupId = groupId;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public Locale getLocale() {
        return locale;
    }

    public String getDescription() {
        return description;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public String getTypeId() {
        return typeId;
    }

    public String getTitle() {
        return title;
    }

    public String getStartDate() {
        return startDate;
    }

    public String isRepeating() {
        return repeating;
    }

    public String getFrequency() {
        return frequency;
    }

    public String getInterval() {
        return interval;
    }

    public String getGroupId() {
        return groupId;
    }

    public String getCenterId() {
        return centerId;
    }

    public void setCenterId(String centerId) {
        this.centerId = centerId;
    }

}
