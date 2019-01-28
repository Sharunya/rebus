package ru.drvsh.rebus;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.MessageFormat;
import java.util.Date;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TestWrightDocx {
    private static final Logger LOGGER = LoggerFactory.getLogger(TestWrightDocx.class);

    public static void main(String[] args) {
        String pathname = null;
        try (XWPFDocument document = new XWPFDocument()) {

            test(document);

            pathname = MessageFormat.format("./test{0}.docx", (new Date()).getTime());
            FileOutputStream fos = new FileOutputStream(new File(pathname));
            document.write(fos);
            fos.close();
        }
        catch (IOException e) {
            LOGGER.error("Ошибка записи", e);
        }
        catch (Exception e) {
            LOGGER.error("Ошибка создания CTPicture", e);
        }
        finally {
            LOGGER.info("Документ word '{}' записан", pathname);
        }
    }

    static void test(XWPFDocument document) {
        // Begin Устанвливаем размер листу и задаем альбомную ориентацию
        CTBody body = document.getDocument().getBody();
        if (!body.isSetSectPr()) {
            body.addNewSectPr();
        }
        CTSectPr section = body.getSectPr();
        if (!section.isSetPgSz()) {
            section.addNewPgSz();
        }
        CTPageSz pageSize = section.getPgSz();

        pageSize.setOrient(STPageOrientation.LANDSCAPE);
        pageSize.setW(BigInteger.valueOf(15840));
        pageSize.setH(BigInteger.valueOf(12240));
        // End У

        /////////

        /// Заголовок (номер и наименование товара)
        XWPFTable tableHeader0 = document.createTable(2, 2);

        XWPFTableRow tableRowOneHeader = tableHeader0.getRow(0);
        XWPFTableCell xwpfTableCell1 = tableRowOneHeader.getCell(0);
        //        CTTbl table = CTTbl.Factory.newInstance();
        //
        //        CTRow r1 = table.addNewTr();
        //        r1.addNewTc().addNewP();

        //        XWPFTable xtab = new XWPFTable(table, document);
        XWPFParagraph paragraph = xwpfTableCell1.getParagraphs().get(0);
        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable tableTwo = xwpfTableCell1.insertNewTbl(cursor);

        XWPFTableRow innerTableRow = tableTwo.createRow();
        XWPFTableCell innerTableRowCell = innerTableRow.createCell();
        innerTableRowCell.setText("qwerty jksdjafio paiosf ai;js dfkasdjuif a[dsj fkjdsaifjudais faslj fkjsdaiofudasi ugfd[is");

        XWPFTableCell xwpfTableCell2 = tableRowOneHeader.getCell(1);
        paragraph = xwpfTableCell2.getParagraphs().get(0);
        cursor = paragraph.getCTP().newCursor();
        tableTwo = xwpfTableCell2.insertNewTbl(cursor);

        innerTableRow = tableTwo.createRow();
        innerTableRowCell = innerTableRow.createCell();
        innerTableRowCell.setText("qwerty jksdjafio paiosf ai;js dfkasdjuif a[dsj fkjdsaifjudais faslj fkjsdaiofudasi ugfd[is");

    }
}
