package ru.drvsh.rebus;

public class ReadExcelDemo {

    public static void main(String[] args) {
        ReaderExcel readerExcel = new ReaderExcel();
        readerExcel.getDataExcel();
        WriterDocx writerDocx = new WriterDocx();
        writerDocx.writeDocx(readerExcel.menuItems, readerExcel.productList, readerExcel.productList.toString());
    }


}
