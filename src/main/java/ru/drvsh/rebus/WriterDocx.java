package ru.drvsh.rebus;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.drvsh.rebus.bean.ProductBeen;
import ru.drvsh.rebus.bean.SpecificationBean;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

public class WriterDocx {
    private static final Logger LOGGER = LoggerFactory.getLogger(WriterDocx.class);

    public void writeDocx(MenuItems menuItems, List<ProductBeen> productList, String fileContent) {
        try (XWPFDocument document = new XWPFDocument()) {

            CTBody body = document.getDocument().getBody();
            if (!body.isSetSectPr()) {
                body.addNewSectPr();
            }
            CTSectPr section = body.getSectPr();
            if(!section.isSetPgSz()) {
                section.addNewPgSz();
            }
            CTPageSz pageSize = section.getPgSz();

            pageSize.setOrient(STPageOrientation.LANDSCAPE);
            pageSize.setW(BigInteger.valueOf(15840));
            pageSize.setH(BigInteger.valueOf(12240));


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
            tableRowOneProduct0.getCell(0).setText("Номер товара в таблице");
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
            tableRowFourProduct0.getCell(2).setText(product0Specification1.getRequirement());
            tableRowFourProduct0.getCell(3).setText(product0Specification2.getRequirement());

            ////
            ////
            FileOutputStream fos;

            fos = new FileOutputStream(new File("./qqqq.docx"));
            document.write(fos);
            fos.close();
        } catch (IOException e) {
            LOGGER.error("Ошибка записи", e);
        } finally {
            LOGGER.info("Документ word записан");
        }
    }
}
