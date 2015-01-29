package nakamurakj.excel.book;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.FileNotFoundException;

import nakamurakj.excel.exception.ExcelIOException;

import org.junit.Test;

public class ExcelTest {

	@Test
	public void readTest() throws ExcelIOException, FileNotFoundException {
		ExcelWorkbook book = ExcelWorkbook.open("template.xls");
		ExcelSheet sheet = book.selectActiveSheet("test");
		assertEquals(14, sheet.getLastIndexRowNumber());

		String account = sheet.getCell(4, "E").getStringValue();
		String package_ = sheet.getCell(5, "E").getStringValue();
		String updated = sheet.getCell(6, "E").getStringValue();

		assertEquals("nakamurakj", account);
		assertEquals("nakamurakj.excel", package_);
		assertEquals("2015/01/22",updated);

		String first = sheet.getCell(9, "C").getStringValue();
		String last = sheet.getCell(9, "D").getStringValue();
		String mailaddress = sheet.getCell(9, "E").getStringValue();
		String pass = sheet.getCell(9, "F").getStringValue();
		String zip = sheet.getCell(9, "G").getStringValue();
		String address1 = sheet.getCell(9, "H").getStringValue();
		String address2 = sheet.getCell(9, "I").getStringValue();
		String address3 = sheet.getCell(9, "J").getStringValue();
		String tel = sheet.getCell(9, "K").getStringValue();
		String comment = sheet.getCell(9, "L").getStringValue();
		String etc = sheet.getCell(9, "M").getStringValue();
		assertEquals("Nobunaga", first);
		assertEquals("Oda", last);
		assertEquals("aaa@test.com", mailaddress);
		assertEquals("111111", pass);
		assertEquals("111-1111", zip);
		assertEquals("aaa", address1);
		assertEquals("bbb", address2);
		assertEquals("ccc", address3);
		assertEquals("111-1111-1111", tel);
		assertEquals("hello", comment);
		assertEquals("", etc);

		first = sheet.getCell(10, "C").getStringValue();
		last = sheet.getCell(10, "D").getStringValue();
		mailaddress = sheet.getCell(10, "E").getStringValue();
		pass = sheet.getCell(10, "F").getStringValue();
		zip = sheet.getCell(10, "G").getStringValue();
		address1 = sheet.getCell(10, "H").getStringValue();
		address2 = sheet.getCell(10, "I").getStringValue();
		address3 = sheet.getCell(10, "J").getStringValue();
		tel = sheet.getCell(10, "K").getStringValue();
		comment = sheet.getCell(10, "L").getStringValue();
		etc = sheet.getCell(10, "M").getStringValue();
		assertEquals("Hideyoshi", first);
		assertEquals("Hashiba", last);
		assertEquals("bbb@test.com", mailaddress);
		assertEquals("222222", pass);
		assertEquals("222-2222", zip);
		assertEquals("ddd", address1);
		assertEquals("eee", address2);
		assertEquals("fff", address3);
		assertEquals("222-2222-2222", tel);
		assertEquals("", comment);
		assertEquals("hi", etc);

		first = sheet.getCell(11, "C").getStringValue();
		if (!first.isEmpty()) {
			fail();
		}
	}
}
