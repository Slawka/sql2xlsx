package ru.ybar.sql2xlsx;

import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

public class sql2xlsx {

    /**
     *
     * @param result sql result
     * @param workbook Workbook
     * @return Workbook
     * @throws IOException
     * @throws SQLException
     */
    public Workbook sql2xlsx(ResultSet result, Workbook workbook) throws IOException, SQLException {
        return sql2xlsx(result, workbook, "1");
    }

    /**
     *
     * @param result sql result
     * @param workbook Workbook
     * @param sheetName String
     * @return Workbook
     * @throws FileNotFoundException
     * @throws IOException
     * @throws SQLException
     */
    public Workbook sql2xlsx(ResultSet result, Workbook workbook, String sheetName) throws FileNotFoundException, IOException, SQLException {

        Sheet sheet = workbook.createSheet(sheetName);
        // Create a Row
        int rowNum = 0;
        Row headerRow = sheet.createRow(rowNum);

        ResultSetMetaData metaData = result.getMetaData();
        int colCount = metaData.getColumnCount();

        for (int i = 0; i < colCount; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(metaData.getColumnLabel(i + 1));
        }
        Cell cell;
        while (result.next()) {
            rowNum++;
            Row row = sheet.createRow(rowNum);
            for (int i = 0; i < colCount; i++) {
                cell = row.createCell(i);
                cell.setCellValue(result.getString(i + 1));
           }
        }
        for (int i = 0; i < colCount; i++) {
            sheet.autoSizeColumn(i);
        }
        return workbook;
    }
}