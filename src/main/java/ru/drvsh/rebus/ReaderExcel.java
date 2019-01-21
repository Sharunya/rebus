package ru.drvsh.rebus;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReaderExcel {
    private File file = new File("/home/lihacheff/Work/out/rebus/./source.xlsx");

    public CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    private XSSFCell getCellByAddress(XSSFSheet sheet, CellAddress address) {
        XSSFRow row = sheet.getRow(address.getRow());
        if (row != null) {
            return row.getCell(address.getColumn());
        }
        return null;
    }

    public void readExcel() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        Iterator<Sheet> iterator = workbook.sheetIterator();
        while (iterator.hasNext()) {
            XSSFSheet sheet = (XSSFSheet) iterator.next();
            System.out.println("\nДанные из таблицы: " + sheet.getSheetName());

            for (int r = 0; r < sheet.getLastRowNum() + 1; r++) {
                StringBuilder data = new StringBuilder();
                XSSFRow row = sheet.getRow(r);
                for (int c = 0; row != null && c < row.getLastCellNum(); c++) {
                    XSSFCell cell = getCell(sheet, row, c);
                    fillValue(data, cell);

                    data.append("\t|\t");
                }
                System.out.println(data);
            }
        }
        workbook.close();
    }

    private XSSFCell getCell(XSSFSheet sheet, XSSFRow row, int cellNum) {
        XSSFCell cell = row.getCell(cellNum);
        CellRangeAddress rangeAddressa = getMergedRegion(sheet, row.getRowNum(), (short) cell.getColumnIndex());
        if(rangeAddressa != null){
            for (CellAddress address : rangeAddressa) {
                XSSFCell cellTmp = getCellByAddress(sheet, address);
                if (cellTmp != null) {
                    cell = cellTmp;
                    break;
                }
            }
        }
        return cell;
    }

    private void fillValue(StringBuilder data, XSSFCell cell) {
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case _NONE:
                break;
            case NUMERIC:
                data.append(cell.getNumericCellValue());
                break;
            case STRING:
                data.append(cell.getStringCellValue());
                break;
            case FORMULA:
                data.append(cell.getCellFormula());
                break;
            case BLANK:
                break;
            case BOOLEAN:
                data.append(cell.getBooleanCellValue());
                break;
            case ERROR:
                data.append("!");
                break;
        }
    }

    public void getDataExcel() {
        try {
            readExcel();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
