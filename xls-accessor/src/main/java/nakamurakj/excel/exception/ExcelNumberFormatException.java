
package nakamurakj.excel.exception;

/**
 * セル操作時の値間違い例外
 */
public class ExcelNumberFormatException extends ExcelIOException {

    /**
     * シリアルID
     */
    private static final long serialVersionUID = -2334347470094174518L;

    /**
     * デフォルトコンストラクタ
     */
    public ExcelNumberFormatException() {
        super();
    }

    /**
     * メッセージ設定のコンストラクタ
     *
     * @param s メッセージ
     */
    public ExcelNumberFormatException(String s) {
        super(s);
    }

}
