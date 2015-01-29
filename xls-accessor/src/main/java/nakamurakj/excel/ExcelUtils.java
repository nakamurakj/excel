
package nakamurakj.excel;

import static nakamurakj.excel.ExcelConst.COLOMN_FOOTS;
import static nakamurakj.excel.ExcelConst.COLOMN_HEADS;
import static nakamurakj.excel.ExcelConst.COLOMN_MAX_NUM;
import static nakamurakj.excel.ExcelConst.COLOMN_MIN_NUM;
import static nakamurakj.excel.ExcelConst.ROW_MAX_NUM;
import static nakamurakj.excel.ExcelConst.ROW_MIN_NUM;
import static org.apache.commons.lang3.StringUtils.isNotEmpty;
import nakamurakj.excel.exception.ExcelNumberFormatException;


/**
 * エクセルユーティリティクラス
 */
public class ExcelUtils {

    /**
     * POI操作で使用する列数を取得する。
     *
     * @param number 開始位置
     * @return Excelの表示で1の箇所は0を返す。
     */
    private static int getNumber(int number) {
        return number - 1;
    }

    /**
     * 行番号を取得する。
     *
     * @param number 表示行番号
     * @return 範囲外の場合は
     */
    public static int getRowNumber(int number) throws ExcelNumberFormatException {
        int ret = getNumber(number);
        if (ret < ROW_MIN_NUM || ROW_MAX_NUM < ret) {
            throw new ExcelNumberFormatException("row number(" + number + ") is outside");
        }
        return ret;
    }

    /**
     * 列文字列から、列番号を取得する。
     *
     * @param colomn 列文字列
     * @return 範囲内の場合は列番号、範囲外の場合は-1
     * @throws ExcelNumberFormatException 列の表記に誤りがある場合
     */
    public static int getColomnNumber(String colomn) throws ExcelNumberFormatException {
        int ret = 0;
        if (isNotEmpty(colomn)) {
            int length = colomn.length();
            if (length == 1) {
                for (int i = 0; i < COLOMN_FOOTS.length; i++) {
                    if (COLOMN_FOOTS[i].equals(colomn)) {
                        ret = i + 1;
                        break;
                    }
                }
            } else if (length == 2) {
                for (int i = 0; i < COLOMN_HEADS.length; i++) {
                    if (colomn.startsWith(COLOMN_HEADS[i])) {
                        for (int j = 0; j < COLOMN_FOOTS.length; i++) {
                            if (colomn.endsWith(COLOMN_FOOTS[j])) {
                                ret = (COLOMN_HEADS.length * i) + j + 1;
                                break;
                            }
                        }
                    }
                }
            }
        }
        ret = getNumber(ret);
        if (ret < COLOMN_MIN_NUM || COLOMN_MAX_NUM < ret) {
            throw new ExcelNumberFormatException("colomn number(" + colomn + ") is outside");
        }
        return ret;
    }
}
