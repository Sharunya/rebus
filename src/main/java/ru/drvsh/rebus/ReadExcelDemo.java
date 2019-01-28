package ru.drvsh.rebus;

import java.text.ParseException;

public class ReadExcelDemo {

    public static void main(String[] args) throws ParseException {
        ReaderExcel readerExcel = new ReaderExcel();
        if(readerExcel.getDataExcel()) {

            WriterDocx writerDocx = new WriterDocx();
            writerDocx.writeDocx(readerExcel.menuItems, readerExcel.productList, readerExcel.selectedFile.getAbsolutePath().split("/./")[0]);
        }
    }

}
