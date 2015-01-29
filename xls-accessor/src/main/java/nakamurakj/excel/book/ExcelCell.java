
package nakamurakj.excel.book;

import java.util.Date;



import nakamurakj.excel.exception.ExcelIOException;

import org.apache.poi.ss.usermodel.Cell;


/**
 * セル操作クラス
 */
public class ExcelCell {

    /** {@link Cell}オブジェクト */
    private Cell cell;

    /** 行番号(表示) */
    private int rowNumber;

    /** 列番号(表示) */
    private String columnNumber;

    /**
     * コンストラクタ
     *
     * @param rowNumber 行番号
     * @param columnNumber 列番号
     * @param cell {@link Cell}オブジェクト
     * @throws ExcelIOException 入出力エラー
     */
    public ExcelCell(int rowNumber, String columnNumber, Cell cell) throws ExcelIOException {
        if (cell == null) {
            throw new ExcelIOException("ExcelCell create error, Cell is null");
        }
        this.rowNumber = rowNumber;
        this.columnNumber = columnNumber;
        this.cell = cell;
    }

    /**
     * {@link Cell}オブジェクトを取得する。
     *
     * @return {@link Cell}オブジェクト
     */
    public Cell getCell() {
        return cell;
    }

    /**
     * 真偽値の値をセルに設定する。
     *
     * @param value 真偽値
     */
    public void setBooleanValue(Boolean value) {
        cell.setCellValue(value);
    }

    /**
     * 真偽値の値をセルから取得する。
     *
     * @return 真偽値
     */
    public Boolean getBooleanValue() {
        return cell.getBooleanCellValue();
    }

    /**
     * 日付型の値をセルに設定する。
     *
     * @param value 日付型
     */
    public void setDateValue(Date value) {
        cell.setCellValue(value);
    }

    /**
     * 日付型の値をセルから取得する。
     *
     * @return 日付型
     */
    public Date getDateValue() {
        return cell.getDateCellValue();
    }

    /**
     * 文字型の値をセルに設定する。
     *
     * @param value 文字型
     */
    public void setStringValue(String value) {
        cell.setCellValue(value);
    }

    /**
     * 文字型の値をセルから取得する。
     *
     * @return 文字型
     */
    public String getStringValue() {
        return cell.getStringCellValue();
    }

    /**
     * 数値型の値をセルに設定する。
     *
     * @param value 数値型
     */
    public void setNumericValue(Double value) {
        cell.setCellValue(value);
    }

    /**
     * 数値型の値をセルから取得する。
     *
     * @return 数値型
     */
    public Double getNumericValue() {
        return cell.getNumericCellValue();
    }

    /**
     * 整数値型の値をセルから取得する。
     *
     * @return 数値型
     */
    public Integer getIntValue() {
        return getNumericValue().intValue();
    }


    /**
     * 行番号(表示)を取得する。
     *
     * @return 行番号(表示)
     */
    public int getRowNumber() {
        return rowNumber;
    }

    /**
     * 列番号(表示)を取得する。
     *
     * @return 列番号(表示)
     */
    public String getColumnNumber() {
        return columnNumber;
    }

}
