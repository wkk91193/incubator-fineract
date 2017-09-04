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
package org.apache.fineract.infrastructure.bulkimport.data;

import java.util.HashMap;
import java.util.Map;

public enum GlobalEntityType {
	

        INVALID(0, "globalEntityType.invalid"),
        CLIENTS(1, "clients"),
        GROUPS(2, "globalEntityType.groups"),
        CENTERS(3, "globalEntityType.centers"),
        OFFICES(4, "globalEntityType.offices"),
        STAFF(5, "globalEntityType.staff"),
        USERS(6, "globalEntityType.users"),
        SMS(7, "globalEntityType.sms"),
        DOCUMENTS(8, "globalEntityType.documents"),
        TEMPLATES(9, "globalEntityType.templates"),
        NOTES(10, "globalEntityType.templates"),
        CALENDAR(11, "globalEntityType.calendar"),
        MEETINGS(12, "globalEntityType.meetings"),
        HOLIDAYS(13, "globalEntityType.holidays"),
        LOANS(14, "globalEntityType.loans"),
        LOAN_PRODUCTS(15, "globalEntityType.loanproducts"),
        LOAN_CHARGES(16, "globalEntityType.loancharges"),
        LOAN_TRANSACTIONS(17, "globalEntityType.loantransactions"),
        GUARANTORS(18, "globalEntityType.guarantors"),
        COLLATERALS(19, "globalEntityType.collaterals"),
        FUNDS(20, "globalEntityType.funds"),
        CURRENCY(21, "globalEntityType.currencies"),
        SAVINGS_ACCOUNT(22, "globalEntityType.savingsaccount"),
        SAVINGS_CHARGES(23, "globalEntityType.savingscharges"),
        SAVINGS_TRANSACTIONS(24, "globalEntityType.savingstransactions"),
        SAVINGS_PRODUCTS(25, "globalEntityType.savingsproducts"),
        GL_JOURNAL_ENTRIES(26, "globalEntityType.gljournalentries"),
        CODE_VALUE(27, "codevalue"),
        CODE(28, "code");

	    private final Integer value;
	    private final String code;

	    private static final Map<Integer, GlobalEntityType> intToEnumMap = new HashMap<>();
	    private static final Map<String, GlobalEntityType> stringToEnumMap = new HashMap<>();
	    private static int minValue;
	    private static int maxValue;
	    
	    static {
	        int i = 0;
	        for (final GlobalEntityType entityType : GlobalEntityType.values()) {
	            if (i == 0) {
	                minValue = entityType.value;
	            }
	            intToEnumMap.put(entityType.value, entityType);
	            stringToEnumMap.put(entityType.code, entityType);
	            if (minValue >= entityType.value) {
	                minValue = entityType.value;
	            }
	            if (maxValue < entityType.value) {
	                maxValue = entityType.value;
	            }
	            i = i + 1;
	        }
	    }

	    private GlobalEntityType(final Integer value, final String code) {
	        this.value = value;
	        this.code = code;
	    }

	    public Integer getValue() {
	        return this.value;
	    }

	    public String getCode() {
	        return this.code;
	    }
	    
	    public static GlobalEntityType fromInt(final int i) {
	        final GlobalEntityType entityType = intToEnumMap.get(Integer.valueOf(i));
	        return entityType;
	    }
	    
	    public static GlobalEntityType fromCode(final String key) {
	        final GlobalEntityType entityType = stringToEnumMap.get(key);
	        return entityType;
	    }
	    
	    @Override
	    public String toString() {
	        return name().toString();
	    }

}
