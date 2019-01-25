package ru.drvsh.rebus;

import com.microsoft.schemas.office.office.CTLock;
import com.microsoft.schemas.office.word.STWrapType;
import com.microsoft.schemas.vml.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;
import ru.drvsh.rebus.bean.ProductBeen;
import ru.drvsh.rebus.bean.SpecificationBean;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.MessageFormat;
import java.util.Date;
import java.util.List;

public class WriterDocx {
    private static final Logger LOGGER = LoggerFactory.getLogger(WriterDocx.class);

    public void writeDocx(MenuItems menuItems, List<ProductBeen> productList, String fileContent) {
        String pathname = null;
        try (XWPFDocument document = new XWPFDocument()) {

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

            ////////// Todo Текстовое поле 2
           /* {
                String box_text = "The TextBox text";

                XWPFParagraph paragraph = document.createParagraph();

                XWPFRun run = paragraph.createRun();

                CTGroup ctGroup = CTGroup.Factory.newInstance();
                CTShape ctShape;
                {
                    ctShape = ctGroup.addNewShape();
                    ctShape.addNewWrap().setType(STWrapType.SQUARE);
                    ctShape.setStyle("position:absolute;mso-position-horizontal:center;margin-top:40pt;width:100pt;height:24pt");
                    CTTxbxContent ctTxbxContent = ctShape.addNewTextbox().addNewTxbxContent();
                    ctTxbxContent.addNewP().addNewR().addNewT().setStringValue("1 " + box_text);

                    Node ctGroupNode = ctGroup.getDomNode();
                    CTPicture ctPicture = CTPicture.Factory.parse(ctGroupNode);
                    CTR cTR = run.getCTR();
                    cTR.addNewPict();
                    cTR.setPictArray(0, ctPicture);
                }

            }*/

            /////////


            /////////
            ProductBeen product = productList.get(0);

            XWPFTable tableHeader0 = document.createTable(2, 2);

            XWPFTableRow tableRowOneHeader0 = tableHeader0.getRow(0);
            tableRowOneHeader0.getCell(0).setText(menuItems.getIdItemName());
            tableRowOneHeader0.getCell(1).setText(menuItems.getNameItemName());

            XWPFTableRow tableRowTwo = tableHeader0.getRow(1);
            tableRowTwo.getCell(0).setText(product.getId());
            tableRowTwo.getCell(1).setText(product.getName());

            ////

            XWPFParagraph tmpParagraph = document.createParagraph();
            XWPFRun tmpRun = tmpParagraph.createRun();
            tmpRun.setText("");
            tmpRun.setFontSize(10);

            ////

            List<SpecificationBean> specificationList = product.getSpecificationList();
            XWPFTable tableProduct0 = document.createTable(4, 4);

            XWPFTableRow tableRowOneProduct0 = tableProduct0.getRow(0);

            /// Todo paragraph CTShape
            XWPFParagraph addParagraph = tableRowOneProduct0.getCell(0).getParagraphs().get(0);

            XWPFRun run = addParagraph.createRun();
            CTGroup ctGroup = CTGroup.Factory.newInstance();

            CTShape ctShape = ctGroup.addNewShape();
            ctShape.addNewWrap().setType(STWrapType.SQUARE);
            ctShape.setStyle("position:inherit;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
            ctShape.setFillcolor("black");
            ctShape.setStroked(STTrueFalse.FALSE);
            CTLock lock = ctShape.addNewLock();
            lock.setExt(STExt.VIEW);
            CTTxbxContent ctTxbxContent = ctShape.addNewTextbox().addNewTxbxContent();
            ctTxbxContent.addNewP().addNewR().addNewT().setStringValue("Номер товара в таблице");

            Node ctGroupNode = ctGroup.getDomNode();
            CTPicture ctPicture = CTPicture.Factory.parse(ctGroupNode);
            CTR cTR = run.getCTR();
            cTR.addNewPict();
            cTR.setPictArray(0, ctPicture);


//            tableRowOneProduct0.getCell(0).setText("Номер товара в таблице");
            tableRowOneProduct0.getCell(1).setText(product.getId());
            tableRowOneProduct0.getCell(2).setText(product.getId());
            tableRowOneProduct0.getCell(3).setText(product.getId());

            SpecificationBean product0Specification0 = specificationList.get(0);
            SpecificationBean product0Specification1 = specificationList.get(1);
            SpecificationBean product0Specification2 = specificationList.get(2);

            XWPFTableRow tableRowTwoProduct0 = tableProduct0.getRow(1);
            tableRowTwoProduct0.getCell(0).setText(menuItems.getSpecificationItemName());
            tableRowTwoProduct0.getCell(1).setText(product0Specification0.getSpecification());
            tableRowTwoProduct0.getCell(2).setText(product0Specification1.getSpecification());
            tableRowTwoProduct0.getCell(3).setText(product0Specification2.getSpecification());

            XWPFTableRow tableRowThreeProduct0 = tableProduct0.getRow(2);
            tableRowThreeProduct0.getCell(0).setText(menuItems.getIdClauseItemName());
            tableRowThreeProduct0.getCell(1).setText(product0Specification0.getClause());
            tableRowThreeProduct0.getCell(2).setText(product0Specification1.getClause());
            tableRowThreeProduct0.getCell(3).setText(product0Specification2.getClause());

            XWPFTableRow tableRowFourProduct0 = tableProduct0.getRow(3);
            tableRowFourProduct0.getCell(0).setText(menuItems.getRequirementItemName());
            tableRowFourProduct0.getCell(1).setText(product0Specification0.getRequirement());
            /////// todo!
            tableRowFourProduct0.getCell(2).setParagraph(textBox(document, product0Specification1.getRequirement()));
//            tableRowFourProduct0.getCell(2).setText(product0Specification1.getRequirement());
            tableRowFourProduct0.getCell(3).setText(product0Specification2.getRequirement());

            ////

            ////
            FileOutputStream fos;

            pathname = MessageFormat.format("./qqqq{0}.docx", (new Date()).getTime());
            fos = new FileOutputStream(new File(pathname));
            document.write(fos);
            fos.close();
        } catch (IOException e) {
            LOGGER.error("Ошибка записи", e);
        } catch (XmlException e) {
            LOGGER.error("Ошибка создания CTPicture", e);
        } finally {
            LOGGER.info("Документ word '{}' записан", pathname);
        }
    }


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
        shape.setId("PowerPlusWaterMarkObject"/* + idx*/);
        shape.setSpid("_x0000_s102" + (4/* + idx*/));
        shape.setType("#_x0000_t136");
        shape.setStyle("position:absolute;margin-left:0;margin-top:0;width:415pt;height:207.5pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin");
        shape.setFillcolor("black");
        shape.setStroked(STTrueFalse.FALSE);
        CTTextPath shapeTextPath = shape.addNewTextpath();
        shapeTextPath.setStyle("font-family:&quot;Cambria&quot;;font-size:1pt");
        shapeTextPath.setString(text);
        pict.set(group);
        // end watermark paragraph
        return new XWPFParagraph(p, document);
    }



    /**
     * Constructs a textbox under the drawing.
     *
     * @param anchor    the client anchor describes how this group is attached
     *                  to the sheet.
     * @return      the newly created textbox.
     */
    public XSSFTextBox createTextbox(XSSFClientAnchor anchor){
        CTTwoCellAnchor ctAnchor = createTwoCellAnchor(anchor);
        CTShape ctShape = ctAnchor.addNewSp();
        ctShape.set(XSSFSimpleShape.prototype());
        XSSFTextBox shape = new XSSFTextBox(this, ctShape);
        shape.anchor = anchor;
        return shape;

    }

}
