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
package org.apache.fineract.infrastructure.bulkimport.importhandler;

import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import org.apache.fineract.infrastructure.bulkimport.constants.TemplatePopulateImportConstants;
import org.apache.poi.ss.usermodel.*;
import org.joda.time.LocalDate;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public class ImportHandlerUtils  {

    public static Integer getNumberOfRows(Sheet sheet, int primaryColumn) {
        Integer noOfEntries = 1;
        // getLastRowNum and getPhysicalNumberOfRows showing false values
        // sometimes
        while (sheet.getRow(noOfEntries) !=null && sheet.getRow(noOfEntries).getCell(primaryColumn) != null) {
            noOfEntries++;
        }

        return noOfEntries;
    }

    public static Boolean isNotImported(Row row, int statusColumn) {
        return !readAsString(statusColumn, row).equals(TemplatePopulateImportConstants.STATUS_CELL_IMPORTED);
    }

    public static Long readAsLong(int colIndex, Row row) {
            Cell c = row.getCell(colIndex);
            if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
                return null;
            FormulaEvaluator eval = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            if(c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                if(eval!=null) {
                    CellValue val = eval.evaluate(c);
                    return ((Double) val.getNumberValue()).longValue();
                }
            }
            else if (c.getCellType()==Cell.CELL_TYPE_NUMERIC){
                return ((Double) c.getNumericCellValue()).longValue();
            }
            else {
                return Long.parseLong(row.getCell(colIndex).getStringCellValue());
            }
            return null;
    }


    public static String readAsString(int colIndex, Row row) {

            Cell c = row.getCell(colIndex);
            if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
                return "";
            FormulaEvaluator eval = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            if(c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    if (eval!=null) {
                        CellValue val = eval.evaluate(c);
                        String res = trimEmptyDecimalPortion(val.getStringValue());
                        return res.trim();
                    }
            }else if(c.getCellType()==Cell.CELL_TYPE_STRING) {
                String res = trimEmptyDecimalPortion(c.getStringCellValue().trim());
                return res.trim();

            }else  {
                return ((Double) row.getCell(colIndex).getNumericCellValue()).intValue() + "";
            }
          return null;
    }


    public static String trimEmptyDecimalPortion(String result) {
        if(result != null && result.endsWith(".0"))
            return	result.split("\\.")[0];
        else
            return result;
    }

    public static LocalDate readAsDate(int colIndex, Row row, final String format) {
        try{
            Cell c = row.getCell(colIndex);
            if(c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
                return null;

            DateFormat dateFormat = new SimpleDateFormat(format);
            Date date=dateFormat.parse(dateFormat.format(c.getDateCellValue()));
            LocalDate localDate=new LocalDate(date);
            return localDate;
        }  catch  (ParseException e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void writeString(int colIndex, Row row, String value) {
        if(value!=null)
        row.createCell(colIndex).setCellValue(value);
    }

    public static CellStyle getCellStyle(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return style;
    }

    public static String parseStatus(String errorMessage) {
        StringBuffer message = new StringBuffer();
        JsonObject obj = new JsonParser().parse(errorMessage.trim()).getAsJsonObject();
        JsonArray array = obj.getAsJsonArray("errors");
        Iterator<JsonElement> iterator = array.iterator();
        while(iterator.hasNext()) {
            JsonObject json = iterator.next().getAsJsonObject();
            String parameterName = json.get("parameterName").getAsString();
            String defaultUserMessage = json.get("defaultUserMessage").getAsString();
            message = message.append(parameterName.substring(1, parameterName.length() - 1) + ":" + defaultUserMessage.substring(1, defaultUserMessage.length() - 1) + "\t");
        }
        return message.toString();
    }

}