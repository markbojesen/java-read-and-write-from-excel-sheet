import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

public class ExcelWriter {

    private static String[] columns = {"Name", "Email", "Date of Birth", "Salary"};
    private static List<Employee> employees = new ArrayList<Employee>();

    // Initializing employees data to insert into the excel file.
    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1988, 3, 11);
        employees.add(new Employee("Mark Bojesen", "mark@mail.com", dateOfBirth.getTime(), 18500));

        dateOfBirth.set(1990, 06, 16);
        employees.add(new Employee("Bob Bobberson", "bob@mail.com", dateOfBirth.getTime(), 20000));

        dateOfBirth.set(1965, 05, 05);
        employees.add(new Employee("John Johnson", "john@mail.com", dateOfBirth.getTime(), 12000));

        dateOfBirth.set(1981, 07, 19);
        employees.add(new Employee("Tom Tomson", "tom@mail.com", dateOfBirth.getTime(), 25000));
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // New Workbook
        Workbook workbook = new XSSFWorkbook();

        /* CreationHelper helps us create instances for various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Employee");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a row
        Row headerRow = sheet.createRow(0);

        // Creating cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Cell style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        //Create other rows and cells with emplyees data
        int rowNum = 1;
        for (Employee employee : employees) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(employee.getName());
            row.createCell(1).setCellValue(employee.getEmail());

            Cell dateOfBirth = row.createCell(2);
            dateOfBirth.setCellValue(employee.getDateOfBirth());
            dateOfBirth.setCellStyle(dateCellStyle);

            row.createCell(3).setCellValue(employee.getSalary());
        }

        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();
    }
}
