package org.apache.fineract.infrastructure.bulkimport.populator.office;

import org.apache.fineract.infrastructure.bulkimport.populator.AbstractWorkbookPopulator;
import org.apache.fineract.infrastructure.bulkimport.populator.OfficeSheetPopulator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class OfficeWorkbookPopulator extends AbstractWorkbookPopulator {
	private OfficeSheetPopulator officeSheetPopulator;
	
	

	public OfficeWorkbookPopulator(OfficeSheetPopulator officeSheetPopulator) {
		this.officeSheetPopulator = officeSheetPopulator;
	}



	@Override
	public void populate(Workbook workbook) {
		officeSheetPopulator.populate(workbook);	
	}

}
