package ru.drvsh.rebus.wraps;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class RowWrap<T extends XSSFRow> {
    private static final Logger LOGGER = LoggerFactory.getLogger(RowWrap.class.getName());
    private final T row;

    public RowWrap(Row row) {
        this.row = (T) row;
    }

    public T getRow() {
        return row;
    }

    public int getRowNum() {
        return row.getRowNum();
    }

    public CellWrap getCell(int i) {
        return new CellWrap(row.getCell(i));
    }

    public int getLastCellNum() {
        return row.getLastCellNum();
    }
}
