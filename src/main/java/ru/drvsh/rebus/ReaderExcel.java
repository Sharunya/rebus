package ru.drvsh.rebus;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.drvsh.rebus.bean.ClauseBean;
import ru.drvsh.rebus.bean.ProductBeen;
import ru.drvsh.rebus.wraps.CellWrap;
import ru.drvsh.rebus.wraps.RowWrap;
import ru.drvsh.rebus.wraps.SheetWrap;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReaderExcel {

    private static final Logger LOGGER = LoggerFactory.getLogger(ReaderExcel.class);
    private static final String SHEET_NAME_CODS = "Коды";

    private File file = new File("./source.xlsx");

    private List<ClauseBean> clauseBeanList = new ArrayList<>();


    public void readExcel() throws IOException {
        Iterator<Sheet> iterator;
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            iterator = workbook.sheetIterator();
        }
        while (iterator.hasNext()) {
            SheetWrap sheet = new SheetWrap(iterator.next());
            LOGGER.info("Данные из таблицы: {}", sheet.getSheetName());
            if (SHEET_NAME_CODS.equals(sheet.getSheetName())) {
                for (Row rowIter : sheet.getSheet()) {
                    RowWrap row = new RowWrap(rowIter);
                    clauseBeanList.add(new ClauseBean(row.getCell(0).getStrValue(), row.getCell(1).getStrValue()));
                }
            } else {

                for (Row rowIter : sheet.getSheet()) {
                    StringBuilder data = new StringBuilder();
                    RowWrap row = new RowWrap(rowIter);
                    CellRangeAddress currProdBlock = sheet.getCurrProdBlock(row);
                    ProductBeen product = null;
                    if (currProdBlock != null) {
                        int firstColumn = currProdBlock.getFirstColumn();
                        product = new ProductBeen(row.getCell(firstColumn).getStrValue(), row.getCell(firstColumn+1).getStrValue());

                    }
                    System.out.print(product);
                    for (Cell cellIter : row.getRow()) {
                        CellWrap cell = new CellWrap(cellIter);
//                        XSSFCell celll = sheet.getCellBy(cell);
                        fillValue(data, cell);
                    }
                    LOGGER.info(data.toString());
                }
            }
        }
    }

    private boolean fillValue(StringBuilder data, CellWrap cell) {
        if (cell == null) return false;
        data.append(cell.getStrValue());
        data.append("\t|\t");
        return true;
    }

    public void getDataExcel() {
        try {
            readExcel();
        } catch (IOException e) {
            LOGGER.error("Ошибка чтения", e);
        }
    }

}
