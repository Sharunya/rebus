package ru.drvsh.rebus;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.MessageFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ru.drvsh.rebus.bean.ProductBeen;
import ru.drvsh.rebus.bean.SpecificationBean;

public class WriterDocx {
    private static final Logger LOGGER = LoggerFactory.getLogger(WriterDocx.class);

    public void writeDocx(MenuItems menuItems, List<ProductBeen> productList, String path) {

        String pathname = null;
        try (XWPFDocument document = new XWPFDocument()) {
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

            for (ProductBeen product : productList) {

                /// Заголовок (номер и наименование товара)
                XWPFTable tableHeader0 = document.createTable(2, 2);

                XWPFTableRow tableRowOneHeader = tableHeader0.getRow(0);
                fillCell(tableRowOneHeader.getCell(0), menuItems.getIdItemName(), 2000, XWPFTableCell.XWPFVertAlign.TOP);
                fillCell(tableRowOneHeader.getCell(1), menuItems.getNameItemName());

                XWPFTableRow tableRowTwo = tableHeader0.getRow(1);
                fillCell(tableRowTwo.getCell(0), product.getId());
                fillCell(tableRowTwo.getCell(1), product.getName());

                //// разрыв между заголовком и списком показателей
                XWPFParagraph tmpParagraph = document.createParagraph();
                XWPFRun tmpRun = tmpParagraph.createRun();
                tmpRun.setText(" ");
                tmpRun.setFontSize(20);

                //// список показателей
                List<SpecificationBean> specificationList = product.getSpecificationList();
                int tableSpecCount = (int)Math.ceil((float)specificationList.size() / 3);
                Iterator<SpecificationBean> specIterator = specificationList.iterator();

                for (int i = 0; i < tableSpecCount; i++) {

                    XWPFTable tableSpec = document.createTable(4, 4);

                    XWPFTableRow tableSpecRowOne = tableSpec.getRow(0);
                    // заголовок таблицы спецификаций
                    fillCell(tableSpecRowOne.getCell(0), "Номер товара в таблице");
                    fillCell(tableSpecRowOne.getCell(1), product.getId(), null, XWPFTableCell.XWPFVertAlign.CENTER);
                    fillCell(tableSpecRowOne.getCell(2), product.getId(), null, XWPFTableCell.XWPFVertAlign.CENTER);
                    fillCell(tableSpecRowOne.getCell(3), product.getId(), null, XWPFTableCell.XWPFVertAlign.CENTER);

                    // левая обязательная колонка таблицы спецификаций
                    XWPFTableRow tableSpecRowTwo = tableSpec.getRow(1);
                    XWPFTableRow tableSpecRowThree = tableSpec.getRow(2);
                    XWPFTableRow tableSpecRowFour = tableSpec.getRow(3);
                    fillCell(tableSpecRowTwo.getCell(0), menuItems.getSpecificationItemName());
                    fillCell(tableSpecRowThree.getCell(0), menuItems.getIdClauseItemName());
                    fillCell(tableSpecRowFour.getCell(0), menuItems.getRequirementItemName());

                    /// перебираем спецификации
                    int j = 0;
                    while (j++ <= 2) {
                        if (specIterator.hasNext()) {
                            SpecificationBean specification = specIterator.next();
                            fillCell(tableSpecRowTwo.getCell(j), specification.getSpecification());
                            fillCell(tableSpecRowThree.getCell(j), specification.getClause());
                            fillCell(tableSpecRowFour.getCell(j), specification.getRequirement());
                        }
                        else {
                            CTTcPr tcPr = tableSpecRowTwo.getCell(j).getCTTc().addNewTcPr();
                            CTTcBorders ctBorders = tcPr.addNewTcBorders();
                            CTBorder ctB = ctBorders.addNewRight();
                            ctB.setVal(STBorder.NONE);
                            ctB = ctBorders.addNewBottom();
                            ctB.setVal(STBorder.NONE);
                            tcPr = tableSpecRowThree.getCell(j).getCTTc().addNewTcPr();
                            ctBorders = tcPr.addNewTcBorders();
                            ctB = ctBorders.addNewRight();
                            ctB.setVal(STBorder.NONE);
                            ctB = ctBorders.addNewBottom();
                            ctB.setVal(STBorder.NONE);
                            ctB = ctBorders.addNewTop();
                            ctB.setVal(STBorder.NONE);
                            if (j == 2) {
                                ctB = ctBorders.addNewLeft();
                                ctB.setVal(STBorder.NONE);
                            }
                            tcPr = tableSpecRowFour.getCell(j).getCTTc().addNewTcPr();
                            ctBorders = tcPr.addNewTcBorders();
                            ctB = ctBorders.addNewRight();
                            ctB.setVal(STBorder.NONE);
                            ctB = ctBorders.addNewBottom();
                            ctB.setVal(STBorder.NONE);
                            ctB = ctBorders.addNewTop();
                            ctB.setVal(STBorder.NONE);
                            if (j == 2) {
                                ctB = ctBorders.addNewLeft();
                                ctB.setVal(STBorder.NONE);
                            }
                        }

                    }
                    //// разрыв между таблицами показателей
                    XWPFParagraph breackSpecParagraph = document.createParagraph();
                    XWPFRun breackSpecRun = breackSpecParagraph.createRun();
                    breackSpecRun.setText(" ");
                    breackSpecRun.setFontSize(15);

                    ////

                }
            }

            ////

            pathname = MessageFormat.format("{0}/rebus_{1}.docx", path, (new Date()).toLocaleString().replaceAll("\\D", "_"));
            FileOutputStream fos = new FileOutputStream(new File(pathname));
            document.write(fos);
            fos.close();
        }
        catch (IOException e) {
            LOGGER.error("Ошибка записи", e);
        }
        finally {
            LOGGER.info("Документ word '{}' записан", pathname);
        }
    }

    private void fillCell(XWPFTableCell cell, String value) {
        fillCell(cell, value, null, XWPFTableCell.XWPFVertAlign.TOP);
    }

    private void fillCell(XWPFTableCell cell, String value, Integer with, XWPFTableCell.XWPFVertAlign verticalAlignment) {
        cell.setVerticalAlignment(verticalAlignment);
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable tableTwo = cell.insertNewTbl(cursor);

        XWPFTableRow innerTableRow = tableTwo.createRow();
        XWPFTableCell innerTableRowCell = innerTableRow.createCell();
        innerTableRowCell.setVerticalAlignment(verticalAlignment);
        innerTableRowCell.setText(value);

    }

}
