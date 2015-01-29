
package nakamurakj.excel.book;

import static nakamurakj.excel.ExcelUtils.getRowNumber;

import java.util.HashMap;
import java.util.Map;




import nakamurakj.excel.exception.ExcelIOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * シート操作クラス
 */
public class ExcelSheet {

    /** {@link Sheet}オブジェクト */
    private Sheet sheet;

    /** 行格納マップ */
    private Map<Integer, ExcelRow> rows;

    /**
     * コンストラクタ
     *
     * @param workbook {@link Workbook}オブジェクト
     * @throws ExcelIOException 入出力エラー
     */
    public ExcelSheet(Sheet sheet) throws ExcelIOException {
        if (sheet == null) {
            throw new ExcelIOException("ExcelSheet create error, Sheet is null");
        }
        this.sheet = sheet;
        rows = new HashMap<Integer, ExcelRow>(sheet.getLastRowNum() + 1);
    }

    /**
     * {@link Sheet}オブジェクトを取得する。
     *
     * @return {@link Sheet}オブジェクト
     */
    public Sheet getSheet() {
        return sheet;
    }

    /**
     * {@link ExcelCell}を取得する。
     *
     * @param rowNumber 行番号(1～65535)
     * @param colomnName カラム名(A～IV)
     * @return 指定の位置の{@link ExcelCell}
     * @throws ExcelIOException
     */
    public ExcelCell getCell(int rowNumber, String colomnName) throws ExcelIOException {
        int rNum = getRowNumber(rowNumber);
        ExcelRow row = rows.get(rNum);
        if (row == null) {
            row = new ExcelRow(rowNumber, sheet.getRow(rNum));
            rows.put(rNum, row);
        }
        return row.getCell(colomnName);
    }

    /**
     * 最終編集行を取得する。
     * @return 最終編集行
     */
    public int getLastIndexRowNumber() {
    	return sheet.getLastRowNum() + 1;
    }
}
