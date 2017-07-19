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

public class WeeklyMeeting extends Meeting{

    private final String repeatsOnDay;

    public WeeklyMeeting(String startDate, String repeating, String frequency, String interval, String repeatsOnDay, Integer rowIndex  ) {
        super(startDate, repeating, frequency, interval, rowIndex );
        this.repeatsOnDay = repeatsOnDay;
    }

    public String getRepeatsOnDay() {
        return repeatsOnDay;
    }

}
