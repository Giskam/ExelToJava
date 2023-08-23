import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
 
public class ExcelDataProcessor {

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\GiskaVlad\\Desktop\\р.xls";
        String outputFilePath = "H:\\Work\\Іmprovements\\DataFromSCADForExel.xlsx";

        try {
            // Open the input Excel file (.xls format)
            FileInputStream inputStream = new FileInputStream(inputFilePath);
            Workbook inputWorkbook = new HSSFWorkbook(inputStream);

            // Create a new Excel workbook for output (.xlsx format)
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Impruvments");

            // Iterate over sheets in the input file
            for (int sheetIndex = 0; sheetIndex < inputWorkbook.getNumberOfSheets(); sheetIndex++) {
                Sheet inputSheet = inputWorkbook.getSheetAt(sheetIndex);

                // Find cells in the first column where the first value is "N"
                for (Row row : inputSheet) {
                    Cell cell = row.getCell(0); // First column (index 0)
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (cellValue.startsWith("N")) {
                            // Extract decimal numbers using regular expression
                            List<Double> numbers = extractDecimalNumbers(cellValue);

                            // Calculate the minimum and maximum decimal values
                            double min = Double.MAX_VALUE;
                            double max = Double.MIN_VALUE;
                            for (Double numValue : numbers) {
                                min = Math.min(min, numValue);
                                max = Math.max(max, numValue);
                            }

                            // Write the results to the output sheet
                            Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                            Cell minCell = outputRow.createCell(0);
                            minCell.setCellValue(min);
                            Cell maxCell = outputRow.createCell(1);
                            maxCell.setCellValue(max);
                        }
                    }
                }
            }

            // Write the output workbook to the file
            FileOutputStream outputStream = new FileOutputStream(outputFilePath);
            outputWorkbook.write(outputStream);
            outputWorkbook.close();
            outputStream.close();

            // Close the input workbook
            inputWorkbook.close();

            System.out.println("Results written to the output file successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to extract decimal numbers from a given string
    private static List<Double> extractDecimalNumbers(String input) {
        List<Double> numbers = new ArrayList<>();
        Pattern pattern = Pattern.compile("-?\\d+\\.\\d+");
        Matcher matcher = pattern.matcher(input);
        while (matcher.find()) {
            try {
                double numValue = Double.parseDouble(matcher.group());
                numbers.add(numValue);
            } catch (NumberFormatException ignored) {
                // Ignore non-decimal values
            }
        }
        return numbers;
    }
}