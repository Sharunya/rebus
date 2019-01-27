package ru.drvsh.rebus;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

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

public class ReaderExcel {

    private static final Logger LOGGER = LoggerFactory.getLogger(ReaderExcel.class);
    private static final String SHEET_NAME_CODS = "коды";
    private static final int A = 0;
    private static final int B = 1;
    private static final int C = 2;
    private static final int D = 3;
    private static final int E = 4;
    //
    protected Map<String, ClauseBean> clauseList = new HashMap<>();
    protected List<ProductBeen> productList = new ArrayList<>();
    protected MenuItems menuItems;
    private final File file = new File("./source.xlsx");
    /**
     * Первая колонка
     */
    private int firstCell;

    public void readExcel() throws IOException, ParseException {
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
            if (!SHEET_NAME_CODS.equalsIgnoreCase(sheet.getSheetName())) {
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
                        menuItems = new MenuItems(row.getCellStrValue(firstCell + A),
                                                  row.getCellStrValue(firstCell + B),
                                                  row.getCellStrValue(firstCell + C),
                                                  row.getCellStrValue(firstCell + D),
                                                  row.getCellStrValue(firstCell + E));
                        continue;
                    }

                    if (product != null && product.getRangeAddress().containsRow(row.getRowNum())) {
                        product.add(getSpecificationBean(row));
                    }
                    else {
                        product = new ProductBeen(sheet.getCurrProdBlock(row), row.getCellStrValue(firstCell + A), row.getCellStrValue(firstCell + B));
                        product.add(getSpecificationBean(row));
                        productList.add(product);
                    }

                }
            }
        }
    }

    public void getDataExcel() {
        try {
            readExcel();

        }
        catch (IOException e) {
            LOGGER.error("Ошибка чтения", e);
        }
        catch (ParseException e) {
            LOGGER.error("Ошибка парсинга", e);
        }
    }

    private SpecificationBean getSpecificationBean(RowWrap row) throws ParseException {
        ClauseBean clauseBean = clauseList.get(row.getCellStrValue(firstCell + E));
        System.out.println(row.getRawRow().getRowNum());
        return new SpecificationBean(row.getCellStrValue(firstCell + C), row.getCellStrValue(firstCell + D), clauseBean.getId(), clauseBean.getText());
    }

}
