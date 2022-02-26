package Excel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.io.*;
import java.util.Scanner;
/**
 * Imtiaz Adar
 * Excel Using Java -> Write
 */
public class excel_write {
    final static String sheetName = "Information";
    public static void main(String[] args) throws Exception{
        Scanner scan = new Scanner(System.in);
        // serial
        int[] serial = new int[10];
        for(int i = 0; i < 10; i++){
            serial[i] = i + 1;
        }
        // names
        String[] names = new String[10];
        names[0] = "Imtiaz Adar";
        names[1] = "Habib Khan";
        names[2] = "Mahbub Khan";
        names[3] = "Rahim Uddin";
        names[4] = "Abul Kashem";
        names[5] = "Borshon Kabir";
        names[6] = "Ahsan Sani";
        names[7] = "Siam Ahsan";
        names[8] = "Nahin Ahmed";
        names[9] = "Sanowar Khan";
        // phones
        String[] phones = new String[10];
        phones[0] = "8801979554646";
        phones[1] = "8801394939493";
        phones[2] = "8801283239994";
        phones[3] = "8801949034311";
        phones[4] = "8801743222334";
        phones[5] = "8801684848332";
        phones[6] = "8801882838483";
        phones[7] = "8801583847573";
        phones[8] = "8801482828424";
        phones[9] = "8801384829375";
        // address
        String[] address = new String[10];
        address[0] = "Dhaka";
        address[1] = "Rajshahi";
        address[2] = "Chattogram";
        address[3] = "Cumilla";
        address[4] = "Faridpur";
        address[5] = "Bogura";
        address[6] = "Lakshmipur";
        address[7] = "Chandpur";
        address[8] = "Noakhali";
        address[9] = "Sirajgonj";
        // workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);
        XSSFRow row = sheet.createRow(0);
        // font
        Font font = workbook.createFont();
        font.setFontHeight((short) 750);
        font.setFontName("Calibri");
        font.setColor(IndexedColors.RED.getIndex());
        font.setItalic(false);
        //font.setStrikeout(true);
        // cell style
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderTop(BorderStyle.THICK);
        cellStyle.setBorderBottom(BorderStyle.THICK);
        cellStyle.setBorderLeft(BorderStyle.THICK);
        cellStyle.setBorderRight(BorderStyle.THICK);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFont(font);
        // cell
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);
        Cell cell4 = row.createCell(3);
        cell1.setCellStyle(cellStyle);
        cell2.setCellStyle(cellStyle);
        cell3.setCellStyle(cellStyle);
        cell4.setCellStyle(cellStyle);
        //cell value
        cell1.setCellValue("Serial No.");
        cell2.setCellValue("Name");
        cell3.setCellValue("Phone");
        cell4.setCellValue("Address");
        for(int i = 0; i < serial.length; i++){
            row = sheet.createRow(i + 1);
            for(int j = 0; j < 4; j++){
                Cell cell = row.createCell(j);
                cell.setCellStyle(cellStyle);
                if(cell.getColumnIndex() == 0){
                    cell.setCellValue(serial[i]);
                }
                else if(cell.getColumnIndex() == 1){
                    cell.setCellValue(names[i]);
                }
                else if(cell.getColumnIndex() == 2){
                    cell.setCellValue(phones[i]);
                }
                else if(cell.getColumnIndex() == 3){
                    cell.setCellValue(address[i]);
                }
            }
        }
        // column auto size
        for(int i = 0; i < 4; i++)
            sheet.autoSizeColumn(i);
        // save file
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save File");
        fileChooser.setSelectedFile(new File("C:\\Users\\imtia\\OneDrive\\Documents\\Excel_Java\\student_information_offcial.xlsx"));
        if(fileChooser.showSaveDialog(null) == JFileChooser.APPROVE_OPTION){
            File location = fileChooser.getSelectedFile();
            try{
                FileOutputStream bw = new FileOutputStream(location);
                workbook.write(bw);
                bw.close();
                System.out.println("File Written Successfully !");
            }
            catch(Exception e){
                System.out.println("Failed !");
            }
        }
    }
}