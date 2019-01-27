package ru.drvsh.rebus;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.MessageFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ru.drvsh.rebus.bean.ProductBeen;
import ru.drvsh.rebus.bean.SpecificationBean;

public class WriterDocx {
    private static final Logger LOGGER = LoggerFactory.getLogger(WriterDocx.class);

    public void writeDocx(MenuItems menuItems, List<ProductBeen> productList, String fileContent) {

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
                fillCell(tableRowOneHeader.getCell(0), menuItems.getIdItemName(), "20%", XWPFTableCell.XWPFVertAlign.TOP);
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
                    while (specIterator.hasNext() && j++ <= 2) {
                        SpecificationBean specification = specIterator.next();
                        fillCell(tableSpecRowTwo.getCell(j), specification.getSpecification());
                        fillCell(tableSpecRowThree.getCell(j), specification.getClause());
                        fillCell(tableSpecRowFour.getCell(j), specification.getRequirement());
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

            pathname = MessageFormat.format("./qqqq{0}.docx", (new Date()).getTime());
            FileOutputStream fos = new FileOutputStream(new File(pathname));
            document.write(fos);
            fos.close();
        }
        catch (IOException e) {
            LOGGER.error("Ошибка записи", e);
        }
        catch (XmlException e) {
            LOGGER.error("Ошибка создания CTPicture", e);
        }
        finally {
            LOGGER.info("Документ word '{}' записан", pathname);
        }
    }

    private void fillCell(XWPFTableCell cell, String value) throws XmlException {
        fillCell(cell, value, null, XWPFTableCell.XWPFVertAlign.TOP);
    }

    private void fillCell(XWPFTableCell cell, String value, String percentWith, XWPFTableCell.XWPFVertAlign verticalAlignment) {
        cell.setText(value);
        cell.setVerticalAlignment(verticalAlignment);
        if (percentWith != null) {
            cell.setWidthType(TableWidthType.PCT);
            cell.setWidth(percentWith);
        }
/*
        XWPFParagraph addParagraph = cell.getParagraphs().get(0);

        XWPFRun run = addParagraph.createRun();

        CTGroup ctGroup = CTGroup.Factory.newInstance();

        CTShape ctShape = ctGroup.addNewShape();
        CTWrap ctWrap = ctShape.addNewWrap();
        ctWrap.setType(STWrapType.SQUARE);
        ctWrap.setSide(STWrapSide.LARGEST);
        ctShape.setStyle(
            "position:absolute;rotation:0;width:182.1pt;height:132.9pt;mso-wrap-distance-left:0pt;mso-wrap-distance-right:0pt;mso-wrap-distance-top:0pt;mso-wrap-distance-bottom:0pt;margin-top:0pt;mso-position-vertical:top;mso-position-vertical-relative:text;margin-left:0pt;mso-position-horizontal-relative:text"
            //                "position:inherit;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
                        );
        ctShape.setFillcolor("black");
        ctShape.setStroked(STTrueFalse.FALSE);
        CTLock lock = ctShape.addNewLock();
        lock.setExt(STExt.VIEW);
        CTTextbox ctTextbox = ctShape.addNewTextbox();
        ctTextbox.setInset("0.06in,0.06in,0.06in,0.06in");
        CTTxbxContent ctTxbxContent = ctTextbox.addNewTxbxContent();
        ctTxbxContent.addNewP().addNewR().addNewT().setStringValue(value);

        Node ctGroupNode = ctGroup.getDomNode();
        CTPicture ctPicture = CTPicture.Factory.parse(ctGroupNode);
        CTR cTR = run.getCTR();
        cTR.addNewPict();
        cTR.setPictArray(0, ctPicture);
*/
    }

    /**
     * Constructs a textbox under the drawing.
     *
     * @param anchor the client anchor describes how this group is attached
     *               to the sheet.
     *
     * @return the newly created textbox.
     */
/*
    public XSSFTextBox createTextbox(XSSFClientAnchor anchor) {
        CTTwoCellAnchor ctAnchor = createTwoCellAnchor(anchor);
        CTShape ctShape = ctAnchor.addNewSp();
        ctShape.set(XSSFSimpleShape.prototype());
        XSSFTextBox shape = new XSSFTextBox(this, ctShape);
        shape.anchor = anchor;
        return shape;

    }
*/

/*
    private XWPFParagraph textBox(XWPFDocument document, String text) {
        CTP p = CTP.Factory.newInstance();
        byte[] rsidr = document.getDocument().getBody().getPArray(0).getRsidR();
        byte[] rsidrdefault = document.getDocument().getBody().getPArray(0).getRsidRDefault();
        p.setRsidP(rsidr);
        p.setRsidRDefault(rsidrdefault);
        CTPPr pPr = p.addNewPPr();
        pPr.addNewPStyle().setVal("Header");
        // start watermark paragraph
        CTR r = p.addNewR();
        CTRPr rPr = r.addNewRPr();
        rPr.addNewNoProof();
        CTPicture pict = r.addNewPict();
        CTGroup group = CTGroup.Factory.newInstance();
        CTShapetype shapetype = group.addNewShapetype();
        CTTextPath shapeTypeTextPath = shapetype.addNewTextpath();
        shapeTypeTextPath.setOn(STTrueFalse.T);
        shapeTypeTextPath.setFitshape(STTrueFalse.T);
        CTHandles handles = shapetype.addNewHandles();
        CTH h = handles.addNewH();
        h.setPosition("#0,bottomRight");
        h.setXrange("6629,14971");
        CTLock lock = shapetype.addNewLock();
        lock.setExt(STExt.VIEW);
        CTShape shape = group.addNewShape();
        shape.setId("PowerPlusWaterMarkObject"*/
    /* + idx*//*
);
        shape.setSpid("_x0000_s102" + (4*/
    /* + idx*//*
));
        shape.setType("#_x0000_t136");
        shape.setStyle(
            "position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
        shape.setFillcolor("black");
        shape.setStroked(STTrueFalse.FALSE);
        CTTextPath shapeTextPath = shape.addNewTextpath();
        shapeTextPath.setStyle("font-family:&quot;Cambria&quot;;font-size:1pt");
        shapeTextPath.setString(text);
        pict.set(group);
        // end watermark paragraph
        return new XWPFParagraph(p, document);
    }
*/

}
