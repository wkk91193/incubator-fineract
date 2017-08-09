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
import org.apache.poi.ss.usermodel.*;
import org.joda.time.LocalDate;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public abstract class AbstractImportHandler implements ImportHandler {

    protected Integer getNumberOfRows(Sheet sheet, int primaryColumn) {
        Integer noOfEntries = 1;
        // getLastRowNum and getPhysicalNumberOfRows showing false values
        // sometimes
        while (sheet.getRow(noOfEntries) !=null && sheet.getRow(noOfEntries).getCell(primaryColumn) != null) {
            noOfEntries++;
        }

        return noOfEntries;
    }

    protected boolean isNotImported(Row row, int statusColumn) {
        return !readAsString(statusColumn, row).equals("Imported");
    }

    protected Long readAsLong(int colIndex, Row row) {
        try {
            Cell c = row.getCell(colIndex);
            if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
                return null;
            FormulaEvaluator eval = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            if(c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                CellValue val = null;
                try {
                    val = eval.evaluate(c);
                } catch (NullPointerException npe) {
                    return null;
                }
                return ((Double) val.getNumberValue()).longValue();
            }
            return ((Double) c.getNumericCellValue()).longValue();
        } catch (RuntimeException re) {
            return Long.parseLong(row.getCell(colIndex).getStringCellValue());
        }
    }


    protected String readAsString(int colIndex, Row row) {
        try {
            Cell c = row.getCell(colIndex);
            if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
                return "";
            FormulaEvaluator eval = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            if(c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                CellValue val = null;
                try {
                    val = eval.evaluate(c);
                } catch(NullPointerException npe) {
                    return "";
                }
                String res = trimEmptyDecimalPortion(val.getStringValue());
                return res.trim();
            }
            String res = trimEmptyDecimalPortion(c.getStringCellValue().trim());
            return res.trim();
        } catch (Exception e) {
            e.printStackTrace();
            return ((Double)row.getCell(colIndex).getNumericCellValue()).intValue() + "";
        }
    }

    private String trimEmptyDecimalPortion(String result) {
        if(result != null && result.endsWith(".0"))
            return	result.split("\\.")[0];
        else
            return result;
    }

    protected LocalDate readAsDate(int colIndex, Row row) {
        try{
            Cell c = row.getCell(colIndex);
            if(c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
                return null;

            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            Date date=dateFormat.parse(dateFormat.format(c.getDateCellValue()));
            LocalDate localDate=new LocalDate(date);
            return localDate;
        }  catch  (Exception e) {
            return null;
        }
    }

    protected void writeString(int colIndex, Row row, String value) {
        row.createCell(colIndex).setCellValue(value);
    }

    protected CellStyle getCellStyle(Workbook workbook, IndexedColors color) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return style;
    }

    protected String parseStatus(String errorMessage) {
        StringBuffer message = new StringBuffer();
        JsonObject obj = new JsonParser().parse(errorMessage.trim()).getAsJsonObject();
        JsonArray array = obj.getAsJsonArray("errors");
        Iterator<JsonElement> iterator = array.iterator();
        while(iterator.hasNext()) {
            JsonObject json = iterator.next().getAsJsonObject();
            String parameterName = json.get("parameterName").toString();
            String defaultUserMessage = json.get("defaultUserMessage").toString();
            message = message.append(parameterName.substring(1, parameterName.length() - 1) + ":" + defaultUserMessage.substring(1, defaultUserMessage.length() - 1) + "\t");
        }
        return message.toString();
    }

}