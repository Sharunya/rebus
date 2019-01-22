package ru.drvsh.rebus.wraps;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class SheetWrap<T extends XSSFSheet> {
    private static final Logger LOGGER = LoggerFactory.getLogger(SheetWrap.class.getName());
    private final T sheet;

    @SuppressWarnings("unchecked")
    public SheetWrap(Sheet sheet) {
        this.sheet = (T) sheet;
    }

    public T getSheet() {
        return sheet;
    }

    /*   public XSSFCell getCellBy(CellWrap cell) {
           RowWrap row = cell.getRow();
           if (cell == null) return null;
           CellRangeAddress rangeAddress = getMergedRegion(row.getRowNum(), (short) cell.getColumnIndex());
           if (rangeAddress != null) {
               for (CellAddress address : rangeAddress) {
                   XSSFCell cellTmp = getCellByAddress(address);
                   if (cellTmp != null) {
                       cell = cellTmp;
                       break;
                   }
               }
           }
           return cell;
       }
   */
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
        CellRangeAddress rangeAddress = sheet.getMergedRegions()
                .stream()
                .filter(cellAddresses -> cellAddresses.isInRange(row.getCell(0).getCell()))
                .findFirst()
                .orElse(null);
        return rangeAddress == null ? null : new CellRangeAddress(rangeAddress.getFirstRow(), rangeAddress.getLastRow(), rangeAddress.getFirstColumn(), row.getLastCellNum());
    }

    public String getSheetName() {
        return sheet.getSheetName();
    }
}
