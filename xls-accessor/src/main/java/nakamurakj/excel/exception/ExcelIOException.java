
package nakamurakj.excel.exception;

import java.io.IOException;


/**
 * 入出力例外
 */
public class ExcelIOException extends IOException {

    /**
     * シリアルID
     */
    private static final long serialVersionUID = -2715357729575390247L;

    /**
     * デフォルトコンストラクタ
     */
    public ExcelIOException() {
        super();
    }

    /**
     * コンストラクタ
     *
     * @param message メッセージ
     * @param cause 詳細
     */
    public ExcelIOException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * コンストラクタ
     *
     * @param message メッセージ
     */
    public ExcelIOException(String message) {
        super(message);
    }

    /**
     * コンストラクタ
     *
     * @param cause 詳細
     */
    public ExcelIOException(Throwable cause) {
        super(cause);
    }

}
