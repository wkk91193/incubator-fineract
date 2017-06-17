package org.apache.fineract.infrastructure.bulkimport.populator;

import java.util.List;

import org.apache.fineract.organisation.monetary.data.CurrencyData;
import org.apache.fineract.portfolio.fund.data.FundData;
import org.apache.fineract.portfolio.paymenttype.data.PaymentTypeData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExtrasSheetPopulator extends AbstractWorkbookPopulator {

	private List<FundData> funds;
	private List<PaymentTypeData> paymentTypes;
	private List<CurrencyData> currencies;

	private static final int FUND_ID_COL = 0;
	private static final int FUND_NAME_COL = 1;
	private static final int PAYMENT_TYPE_ID_COL = 2;
	private static final int PAYMENT_TYPE_NAME_COL = 3;
	private static final int CURRENCY_CODE_COL = 4;
	private static final int CURRENCY_NAME_COL = 5;

	
	public ExtrasSheetPopulator(List<FundData> funds, List<PaymentTypeData> paymentTypes,
			List<CurrencyData> currencies) {
		this.funds = funds;
		this.paymentTypes = paymentTypes;
		this.currencies = currencies;
	}

	@Override
	public void populate(Workbook workbook) {
		int fundRowIndex = 1;
		Sheet extrasSheet = workbook.createSheet("Extras");
		setLayout(extrasSheet);
		for (FundData fund : funds) {
			Row row = extrasSheet.createRow(fundRowIndex++);
			writeLong(FUND_ID_COL, row, fund.getId());
			writeString(FUND_NAME_COL, row, fund.getName());
		}
		int paymentTypeRowIndex = 1;
		for (PaymentTypeData paymentType : paymentTypes) {
			Row row;
			if (paymentTypeRowIndex < fundRowIndex)
				row = extrasSheet.getRow(paymentTypeRowIndex++);
			else
				row = extrasSheet.createRow(paymentTypeRowIndex++);
			writeLong(PAYMENT_TYPE_ID_COL, row, paymentType.getId());
			writeString(PAYMENT_TYPE_NAME_COL, row, paymentType.getName().trim().replaceAll("[ )(]", "_"));
		}
		int currencyCodeRowIndex = 1;
		for (CurrencyData currencies : currencies) {
			Row row;
			if (currencyCodeRowIndex < paymentTypeRowIndex)
				row = extrasSheet.getRow(currencyCodeRowIndex++);
			else
				row = extrasSheet.createRow(currencyCodeRowIndex++);

			writeString(CURRENCY_NAME_COL, row, currencies.getName().trim().replaceAll("[ )(]", "_"));
			writeString(CURRENCY_CODE_COL, row, currencies.code());
		}
		extrasSheet.protectSheet("");

	}

	private void setLayout(Sheet worksheet) {
		worksheet.setColumnWidth(FUND_ID_COL, 4000);
		worksheet.setColumnWidth(FUND_NAME_COL, 7000);
		worksheet.setColumnWidth(PAYMENT_TYPE_ID_COL, 4000);
		worksheet.setColumnWidth(PAYMENT_TYPE_NAME_COL, 7000);
		worksheet.setColumnWidth(CURRENCY_NAME_COL, 7000);
		worksheet.setColumnWidth(CURRENCY_CODE_COL, 7000);
		Row rowHeader = worksheet.createRow(0);
		rowHeader.setHeight((short) 500);
		writeString(FUND_ID_COL, rowHeader, "Fund ID");
		writeString(FUND_NAME_COL, rowHeader, "Name");
		writeString(PAYMENT_TYPE_ID_COL, rowHeader, "Payment Type ID");
		writeString(PAYMENT_TYPE_NAME_COL, rowHeader, "Payment Type Name");
		writeString(CURRENCY_NAME_COL, rowHeader, "Currency Type ");
		writeString(CURRENCY_CODE_COL, rowHeader, "Currency Code ");
	}
	public Integer getFundsSize() {
		return funds.size();
	}
	public Integer getPaymentTypesSize() {
		return paymentTypes.size();
	}
	
}
