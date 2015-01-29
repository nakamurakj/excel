
package nakamurakj.excel.book;

import static nakamurakj.excel.ExcelUtils.getColomnNumber;
import nakamurakj.excel.exception.ExcelIOException;

import org.apache.poi.ss.usermodel.Row;


/**
 * 行操作クラス
 */
public class ExcelRow {

    /** {@link Row}オブジェクト */
    private Row row;

    /** 行番号(表示) */
    private int rowNumber;

    /**
     * コンストラクタ
     *
     * @param rowNumber 行番号
     * @param row 行オブジェクト
     * @throws ExcelIOException 入出力エラー
     */
    public ExcelRow(int rowNumber, Row row) throws ExcelIOException {
        if (row == null) {
            throw new ExcelIOException("ExcelRow create error, Row is null");
        }
        this.rowNumber = rowNumber;
        this.row = row;
    }

    /**
     * {@link Row}オブジェクトを取得する。
     *
     * @return {@link Row}オブジェクト
     */
    public Row getRow() {
        return row;
    }

    /**
     * {@link ExcelCell}を取得する。
     *
     * @param colomnName カラム名(A～IV)
     * @return 指定の位置の{@link ExcelCell}
     * @throws ExcelIOException
     */
    public ExcelCell getCell(String colomnName) throws ExcelIOException {
        try {
            return new ExcelCell(rowNumber, colomnName, row.getCell(getColomnNumber(colomnName)));
        } catch (ExcelIOException e) {
            return createCell(colomnName);
        }
    }

    /**
     * {@link ExcelCell}を取得する。
     *
     * @param colomnName カラム名(A～IV)
     * @return 指定の位置の{@link ExcelCell}
     * @throws ExcelIOException
     */
    public ExcelCell createCell(String colomnName) throws ExcelIOException {
        return new ExcelCell(rowNumber, colomnName, row.createCell(getColomnNumber(colomnName)));
    }

    /**
     * 行番号(表示)を取得する。
     *
     * @return 行番号(表示)
     */
    public int getRowNumber() {
        return rowNumber;
    }
}
