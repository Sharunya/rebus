package ru.drvsh.rebus.wraps;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class CellWrap<T extends XSSFCell> {
    private static final Logger LOGGER = LoggerFactory.getLogger(CellWrap.class.getName());
    private final T cell;

    @SuppressWarnings("unchecked")
    public CellWrap(Cell cell) {
        this.cell = (T) cell;
    }

    public T getRawCell() {
        return cell;
    }

    public RowWrap getRow() {
        return new RowWrap(cell.getRow());
    }

    public boolean isNull() {
        return cell == null || getRawValue() == null;
    }

    public String getStrValue() {
        if (this.isNull()) return "";
        String result = null;
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case _NONE:
                break;
            case NUMERIC:
                result = String.valueOf(Math.round(cell.getNumericCellValue()));
                break;
            case STRING:
                result = cell.getStringCellValue();
                break;
            case BOOLEAN:
                result = String.valueOf(cell.getBooleanCellValue());
                break;
            case ERROR:
                result = "!";
                break;
            default:
                result = "";
        }
        return result;
    }

    public String getRawValue() {
        return cell.getRawValue();
    }
}
