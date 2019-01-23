package ru.drvsh.rebus.wraps;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class SheetWrap<T extends XSSFSheet> {
    private static final Logger LOGGER = LoggerFactory.getLogger(SheetWrap.class.getName());
    private final T sheet;
    private List<CellRangeAddress> mergedRegions;

    @SuppressWarnings("unchecked")
    public SheetWrap(Sheet sheet) {
        this.sheet = (T) sheet;
    }

    public List<CellRangeAddress> getMergedRegions() {
        if (mergedRegions == null) mergedRegions = sheet.getMergedRegions();
        return mergedRegions;
    }

    public T getRawSheet() {
        return sheet;
    }

    private CellRangeAddress getMergedRegion(int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    private XSSFCell getCellByAddress(CellAddress address) {
        XSSFRow row = sheet.getRow(address.getRow());
        if (row != null) {
            return row.getCell(address.getColumn());
        }
        return null;
    }

    public CellRangeAddress getCurrProdBlock(RowWrap row) {
        CellRangeAddress rangeAddress = getMergedRegions()
                .stream()
                .filter(cellAddresses -> cellAddresses.isInRange(row.getCell(row.getFirstCellNum()).getRawCell()))
                .findFirst()
                .orElse(null);
        return rangeAddress != null ? new CellRangeAddress(rangeAddress.getFirstRow(), rangeAddress.getLastRow(), rangeAddress.getFirstColumn(), row.getLastCellNum())
                : new CellRangeAddress(row.getRowNum(), row.getRowNum(), row.getFirstCellNum(), row.getLastCellNum());
    }

    public String getSheetName() {
        return sheet.getSheetName();
    }
}
