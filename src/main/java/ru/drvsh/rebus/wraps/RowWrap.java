package ru.drvsh.rebus.wraps;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.stream.IntStream;

public class RowWrap<T extends XSSFRow> {
    private static final Logger LOGGER = LoggerFactory.getLogger(RowWrap.class.getName());
    private final T row;

    @SuppressWarnings("unchecked")
    public RowWrap(Row row) {
        this.row = (T) row;
    }

    public T getRawRow() {
        return row;
    }

    public boolean isEmpty() {
        return IntStream.range(getFirstCellNum(), getLastCellNum()).allMatch(i -> getCell(i).isNull());
    }

    public int getRowNum() {
        return row.getRowNum();
    }

    public CellWrap getCell(int i) {
        return new CellWrap(row.getCell(i));
    }

    public String getCellStrValue(int i) {
        return getCell(i).getStrValue();
    }

    public int getFirstCellNum() {
        return row.getFirstCellNum();
    }

    public int getLastCellNum() {
        return row.getLastCellNum();
    }
}
