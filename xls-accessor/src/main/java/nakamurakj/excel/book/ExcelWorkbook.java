
package nakamurakj.excel.book;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;



import nakamurakj.excel.exception.ExcelIOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


/**
 * ワークブック操作クラス。
 */
public class ExcelWorkbook {

    /** ワークブック */
    private Workbook workbook;

    /**
     * プライベートコンストラクタ
     * 外からアクセスさせないため。
     */
    private ExcelWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }


    /**
     * 既存のワークブックを開く。
     *
     * @param fileName ファイル名
     * @return {@link Workbook}オブジェクト
     * @throws ExcelIOException 入出力エラー
     * @throws FileNotFoundException
     */
    public static ExcelWorkbook open(File file) throws ExcelIOException, FileNotFoundException {
        ExcelWorkbook excelWorkBook = open(new FileInputStream(file));
        return excelWorkBook;
    }

    /**
     * 既存のワークブックを開く。
     *
     * @param fileName ファイル名
     * @return {@link Workbook}オブジェクト
     * @throws ExcelIOException 入出力エラー
     * @throws FileNotFoundException
     */
    public static ExcelWorkbook open(String fileName) throws ExcelIOException, FileNotFoundException {
        String filePath = ExcelWorkbook.class.getClassLoader().getResource(fileName).getFile();
        ExcelWorkbook excelWorkBook = open(new FileInputStream(filePath));
        return excelWorkBook;
    }

    /**
     * 既存のワークブックを開く。
     *
     * @param fullPath ファイル名
     * @return {@link Workbook}オブジェクト
     * @throws ExcelIOException 入出力エラー
     * @throws FileNotFoundException
     */
    public static ExcelWorkbook openFullPath(String fullPath) throws ExcelIOException, FileNotFoundException {
        return open(new File(fullPath));
    }

    /**
     * 既存のワークブックを開く。
     *
     * @param fileName ファイル名
     * @return {@link Workbook}オブジェクト
     * @throws ExcelIOException 入出力エラー
     */
    public static ExcelWorkbook open(InputStream is) throws ExcelIOException {
        ExcelWorkbook excelWorkBook = null;
        try {
            Workbook workbook = WorkbookFactory.create(is);
            excelWorkBook = new ExcelWorkbook(workbook);
        } catch (IOException e) {
            throw new ExcelIOException("I/O error!", e);
        } catch (InvalidFormatException e) {
            throw new ExcelIOException("invalid format at xls or xlsx ", e);
        } finally {
            try {
                is.close();
            } catch (IOException e) {
                throw new ExcelIOException(e);
            }
        }
        return excelWorkBook;
    }


    /**
     * {@link Workbook}をファイルに出力する。
     *
     * @param fileName ファイル名
     * @throws ExcelIOException 入出力エラー
     */
    public void write(String fileName) throws ExcelIOException {
        FileOutputStream out = null;
        File file = new File(fileName);
        try {

            out = new FileOutputStream(file);
            workbook.write(out);
        } catch (IOException e) {
            throw new ExcelIOException(file.getAbsolutePath() + " I/O error!", e);
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                throw new ExcelIOException(e);
            }
        }
    }

    /**
     * 指定のシートをアクティブにして先頭表示に選択する。
     *
     * @param sheetName シート名
     * @return {@link Sheet}オブジェクト
     * @throws ExcelIOException 入出力エラー
     */
    public ExcelSheet selectActiveSheet(String sheetName) throws ExcelIOException {
        try {
            int index = workbook.getSheetIndex(sheetName);
            workbook.setActiveSheet(index);
            workbook.setFirstVisibleTab(index);
            return new ExcelSheet(workbook.getSheetAt(index));
        } catch (Exception e) {
            throw new ExcelIOException(sheetName + " is not exist book.", e);
        }
    }

    /**
     * 指定のシートをコピーし、シート名をリネームする
     *
     * @param sheetNameTo コピー元のシート名
     * @param sheetNameFrom コピー先のシート名
     * @throws ExcelIOException 入出力エラー
     */
    public void cloneAndRenameSheet(String sheetNameTo, String sheetNameFrom)
        throws ExcelIOException {
        try {
            int index = workbook.getSheetIndex(sheetNameTo);
            workbook.cloneSheet(index);
            workbook.setSheetName(workbook.getNumberOfSheets() - 1, sheetNameFrom);
        } catch (Exception e) {
            throw new ExcelIOException(sheetNameTo + " is not exist book.", e);
        }
    }

    /**
     * ワークブックを取得する。
     *
     * @return ワークブック
     */
    public Workbook getWorkbook() {
        return workbook;
    }
}
