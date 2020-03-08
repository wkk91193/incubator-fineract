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
package org.apache.fineract.interoperation.handler;

import static org.apache.fineract.interoperation.util.InteropUtil.ENTITY_NAME_IDENTIFIER;

import com.google.common.base.Strings;
import org.apache.fineract.commands.annotation.CommandType;
import org.apache.fineract.commands.handler.NewCommandSourceHandler;
import org.apache.fineract.infrastructure.core.api.JsonCommand;
import org.apache.fineract.infrastructure.core.data.CommandProcessingResult;
import org.apache.fineract.interoperation.domain.InteropIdentifierType;
import org.apache.fineract.interoperation.service.InteropService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.StringUtils;

@Service
@CommandType(entity = ENTITY_NAME_IDENTIFIER, action = "DELETE")
public class DeleteInteropIdentifierHandler implements NewCommandSourceHandler {

    private final InteropService interopService;

    @Autowired
    public DeleteInteropIdentifierHandler(InteropService interopService) {
        this.interopService = interopService;
    }

    @Transactional
    @Override
    public CommandProcessingResult processCommand(final JsonCommand command) {
        String[] split = command.getUrl().split("/");
        int length = split.length;
        String subIdOrType = Strings.emptyToNull(StringUtils.trimWhitespace(split[length - 1]));
        String idValue = split[length - 2];
        InteropIdentifierType idType = InteropIdentifierType.valueOf(split[length - 3].toUpperCase());
        return this.interopService.deleteAccountIdentifier(idType, idValue, subIdOrType);
    }
}