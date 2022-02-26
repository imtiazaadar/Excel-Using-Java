package Excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
/**
 * Imtiaz Adar
 * Excel Using Java -> Read
 */
public class excel_read {
    public static void main(String[] args) throws Exception {
        ArrayList<String> header = new ArrayList<>();
        ArrayList<Integer> serial = new ArrayList<>();
        ArrayList<String> name = new ArrayList<>();
        ArrayList<String> phone = new ArrayList<>();
        ArrayList<String> address = new ArrayList<>();
        // filter
        JFileChooser openfile = new JFileChooser();
        openfile.setDialogTitle("Open File");
        openfile.removeChoosableFileFilter(openfile.getFileFilter());
        FileNameExtensionFilter fileFilter = new FileNameExtensionFilter("Excel File [XLSX]", "xlsx");
        openfile.setFileFilter(fileFilter);
        if(openfile.showSaveDialog(null) == JFileChooser.APPROVE_OPTION){
            File input = openfile.getSelectedFile();
            try{
                FileInputStream br = new FileInputStream(input);
                XSSFWorkbook workbook = new XSSFWorkbook(br);
                XSSFSheet sheet = workbook.getSheetAt(0);
                Iterator<Row> row = sheet.iterator();
                while(row.hasNext()){
                    Row row1 = row.next();
                    Iterator<Cell> col = row1.cellIterator();
                    while(col.hasNext()){
                        Cell col1 = col.next();
                        if(row1.getRowNum() == 0){
                            header.add(col1.getStringCellValue());
                        }
                        else{
                            if(col1.getColumnIndex() == 0){
                                serial.add((int)col1.getNumericCellValue());
                            }
                            else if(col1.getColumnIndex() == 1){
                                name.add(col1.getStringCellValue());
                            }
                            else if(col1.getColumnIndex() == 2){
                                phone.add(col1.getStringCellValue());
                            }
                            else if(col1.getColumnIndex() == 3){
                                address.add(col1.getStringCellValue());
                            }
                        }
                    }
                }
                br.close();
                // print
                System.out.println("Read Successfully !");
                for(int i = 0; i < header.size(); i++) {
                    System.out.print(header.get(i));
                    if(i < header.size() - 1)
                        System.out.print(" | ");
                }
                System.out.println();
                for(int i = 0; i < phone.size(); i++){
                    System.out.println(serial.get(i) + " | " + name.get(i) + " | " + phone.get(i) + " | " + address.get(i));

                }
                System.out.println("Done !");
            }
            catch(Exception e){
                System.out.println("Failed !");
            }
        }
    }
}
