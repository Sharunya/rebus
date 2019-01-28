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

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
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
    public File selectedFile = new File("./source.xlsx");
    //
    protected Map<String, ClauseBean> clauseList = new HashMap<>();
    protected List<ProductBeen> productList = new ArrayList<>();
    protected MenuItems menuItems;
    /**
     * Первая колонка
     */
    private int firstCell;

    public void readExcel(File selectedFile) throws IOException, ParseException {
        Iterator<Sheet> iterator = null;
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(selectedFile))) {
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

    public boolean getDataExcel() throws ParseException {
        JFileChooser jFileChooser = new JFileChooser();
        jFileChooser.setCurrentDirectory(new File("./"));
        jFileChooser.addChoosableFileFilter(new XlsxFileFilter());
        boolean boo = false;

        do {
            int result = jFileChooser.showOpenDialog(new JFrame());

            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = jFileChooser.getSelectedFile();

                try {
                    readExcel(selectedFile);
                    boo = true;
                }
                catch (IOException e) {
                    LOGGER.error("Ошибка чтения", e);
                    boo = false;
                }
                catch (NotOfficeXmlFileException | POIXMLException e) {
                    LOGGER.error(e.getMessage(), e);
                    boo = false;
                }
            }
            else if (result == JFileChooser.CANCEL_OPTION) {
                return false;
            }

        }
        while (!boo);
        return true;
    }

    private SpecificationBean getSpecificationBean(RowWrap row) throws ParseException {
        ClauseBean clauseBean = clauseList.get(row.getCellStrValue(firstCell + E));
        System.out.println(row.getRawRow().getRowNum());
        return new SpecificationBean(row.getCellStrValue(firstCell + C), row.getCellStrValue(firstCell + D), clauseBean.getId(), clauseBean.getText());
    }

    class XlsxFileFilter extends FileFilter {
        @Override
        public boolean accept(File f) {

            if (f != null) {

                String name = f.getName();

                int i = name.lastIndexOf('.');

                if (i > 0 && i < name.length() - 1)

                {
                    return name.substring(i + 1).equalsIgnoreCase("xlsx");
                }

            }

            return false;

        }

        @Override
        public String getDescription() {
            return "Файлы xlsx";

        }
    }
}
