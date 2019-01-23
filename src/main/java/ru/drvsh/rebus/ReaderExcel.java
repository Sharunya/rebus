package ru.drvsh.rebus;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.drvsh.rebus.bean.ClauseBean;
import ru.drvsh.rebus.bean.ProductBeen;
import ru.drvsh.rebus.bean.SpecificationBean;
import ru.drvsh.rebus.wraps.RowWrap;
import ru.drvsh.rebus.wraps.SheetWrap;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ReaderExcel {

    private static final Logger LOGGER = LoggerFactory.getLogger(ReaderExcel.class);
    private static final String SHEET_NAME_CODS = "Коды";
    private static final int A = 0;
    private static final int B = 1;
    private static final int C = 2;
    private static final int D = 3;
    private static final int E = 4;


    private File file = new File("./source.xlsx");

    private MenuItems menuItems = null;
    private Map<String, ClauseBean> clauseList = new HashMap<>();
    private List<ProductBeen> productList = new ArrayList<>();
    /**
     * Первая колонка
     */
    private int firstCell;


    public void readExcel() throws IOException {
        Iterator<Sheet> iterator;
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            iterator = workbook.sheetIterator();
            SheetWrap sheet = new SheetWrap(workbook.getSheet(SHEET_NAME_CODS));
            LOGGER.info("Данные из таблицы: {}", sheet.getSheetName());
            for (Row rowIter : sheet.getRawSheet()) {
                RowWrap row = new RowWrap(rowIter);
                firstCell = row.getFirstCellNum();
                clauseList.put(row.getCellStrValue(firstCell + A), new ClauseBean(row.getCellStrValue(firstCell + A), row.getCellStrValue(firstCell + B)));
            }
        }

        while (iterator.hasNext()) {
            SheetWrap sheet = new SheetWrap(iterator.next());
            if (!SHEET_NAME_CODS.equals(sheet.getSheetName())) {
                LOGGER.info("Данные из таблицы: {}", sheet.getSheetName());
                ProductBeen product = null;

                for (Row rowIter : sheet.getRawSheet()) {

                    RowWrap row = new RowWrap(rowIter);
                    // пропускаем пустую строку
                    if (row.isEmpty()) {
                        continue;
                    }
                    // берем заголовки
                    if (menuItems == null) {
                        menuItems = new MenuItems(
                                row.getCellStrValue(firstCell + A),
                                row.getCellStrValue(firstCell + B),
                                row.getCellStrValue(firstCell + C),
                                row.getCellStrValue(firstCell + D),
                                row.getCellStrValue(firstCell + E)
                        );
                        continue;
                    }


                    if (product != null && product.getRangeAddress().containsRow(row.getRowNum())) {
                        ClauseBean clauseBean = clauseList.get(row.getCellStrValue(firstCell + E));
                        product.add(new SpecificationBean(
                                row.getCellStrValue(firstCell + C),
                                row.getCellStrValue(firstCell + D),
                                clauseBean.getId(),
                                clauseBean.getText())
                        );
                    } else {
                        product = new ProductBeen(
                                sheet.getCurrProdBlock(row),
                                row.getCellStrValue(firstCell + A),
                                row.getCellStrValue(firstCell + B)
                        );
                        productList.add(product);
                    }

                }
            }
        }
    }

    public void getDataExcel() {
        try {
            readExcel();
        } catch (IOException e) {
            LOGGER.error("Ошибка чтения", e);
        }
    }

}
