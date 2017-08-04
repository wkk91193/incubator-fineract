package org.apache.fineract.infrastructure.bulkimport.data;

import org.apache.fineract.infrastructure.core.exception.GeneralPlatformDomainRuleException;

public enum ImportFormatType {
    
    XLSX ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    XLS ("application/vnd.ms-excel"),
    ODS ("application/vnd.oasis.opendocument.spreadsheet");
    
    
    private final String format;

    private ImportFormatType(String format) {
        this.format= format;
    }

    public String getFormat() {
        return format;
    }
    
    public static ImportFormatType of(String name) {
        for(ImportFormatType type : ImportFormatType.values()) {
            if(type.name().equalsIgnoreCase(name)) {
                return type;
            }
        }
        throw new GeneralPlatformDomainRuleException("error.msg.invalid.file.extension",
        		"Uploaded file extension is not recognized.");
    }
}
