package ru.drvsh.rebus;

import java.text.ParseException;

public class ReadExcelDemo {

    public static void main(String[] args) throws ParseException {
        ReaderExcel readerExcel = new ReaderExcel();
        try {

            if (readerExcel.getDataExcel()) {

                WriterDocx writerDocx = new WriterDocx();
                writerDocx.writeDocx(readerExcel.menuItems, readerExcel.productList, readerExcel.selectedFile.getParentFile().getAbsolutePath());
            }
            System.exit(0);

        } catch (Throwable t) {
            System.exit(1);

        }

    }

}
