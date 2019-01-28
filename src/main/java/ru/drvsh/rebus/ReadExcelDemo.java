package ru.drvsh.rebus;

public class ReadExcelDemo {

    public static void main(String[] args) {
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
