package com.hotcocoacup.mobiletools.xlstoresouces;

import com.google.gson.Gson;
import com.google.gson.JsonIOException;
import com.google.gson.JsonSyntaxException;
import com.google.gson.reflect.TypeToken;
import com.hotcocoacup.mobiletools.xlstoresouces.model.ResEntry;
import com.hotcocoacup.mobiletools.xlstoresouces.model.ResFileEntry;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created by a556679 on 16/02/2016.
 */
public class ResourcesToXls {

    public static final String LOGGER_NAME = "XlsToResources";
    private static Logger logger = Logger.getLogger(LOGGER_NAME);

    private ResEntry parseConfigFile(String configFileName) {
        File file = new File(configFileName);
        Gson gson = new Gson();

        try {
            return gson.fromJson(new FileReader(file), new TypeToken<ResEntry>() {
            }.getType());
        } catch (JsonIOException e) {
            logger.log(Level.SEVERE, "Cannot parse the configuration file", e);
        } catch (JsonSyntaxException e) {
            logger.log(Level.SEVERE, "Cannot parse the configuration file", e);
        } catch (FileNotFoundException e) {
            logger.log(Level.SEVERE, "The configuration file does not exist", e);
        }

        return null;
    }

    public void generateExcel(String configFileName) {
        ResEntry resEntry = parseConfigFile(configFileName);

        if (resEntry == null)
            return;

        // Create the output document. Create & set sheet name
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet spreadSheet = workbook.createSheet(resEntry.getSheetName());

        // Set column width
        int columNb = resEntry.getResourcesFiles().size() + 1;
        spreadSheet.setColumnWidth(0, (256 * 40));
        for (int i = 1; i < columNb; i++) {
            spreadSheet.setColumnWidth(i, (256 * 100));
        }

        HSSFFont font = workbook.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);

        CellStyle cellStyleTitle = workbook.createCellStyle();
        cellStyleTitle.setFont(font);

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(true);

        // Creating first row with style
        HSSFRow firstRow = spreadSheet.createRow(0);
        HSSFCell cell = firstRow.createCell(0);
        cell.setCellValue(resEntry.getFirstColumnName());
        cell.setCellStyle(cellStyleTitle);

        TreeMap<String, List<String>> map = new TreeMap<String, List<String>>();

        // Loop on all XML resources files
        int columnIndex = 1;
        for (ResFileEntry resFileEntry : resEntry.getResourcesFiles()) {

            Document doc = null;

            try {
                doc = DocumentBuilderFactory.newInstance()
                        .newDocumentBuilder()
                        .parse(new File(resFileEntry.getFileName()));
            } catch (SAXException e) {
                logger.log(Level.SEVERE, "XML parsing errors", e);
            } catch (IOException e) {
                logger.log(Level.SEVERE, "IO errors occurs", e);
            } catch (ParserConfigurationException e) {
                logger.log(Level.SEVERE, "DocumentBuilder cannot be created which satisfies the configuration requested", e);
            }

            if (doc == null)
                continue;

            NodeList nList = doc.getElementsByTagName("string");

            // Set column title
            cell = firstRow.createCell(columnIndex);
            cell.setCellValue(resFileEntry.getColumnName());
            cell.setCellStyle(cellStyleTitle);

            System.out.println("\n\n");
            System.out.println("******************************************************************************");
            System.out.println("***************** Column name = " + resFileEntry.getColumnName());
            System.out.println("***************** Column number = " + columnIndex);
            System.out.println("******************************************************************************");

            String key;
            for (int temp = 0; temp < nList.getLength(); temp++) {
                Element eElement = (Element) nList.item(temp);

                key = eElement.getAttribute("name");
                String t = parseValue(eElement.getTextContent());

                // Log pattern : [column ; Line] Keyword : Value
                System.out.println("[" + columnIndex + ";" + (temp + 1) + "] " +
                        key + ":" + new HSSFRichTextString(t));

                if (map.containsKey(key)) {
                    map.get(key).add(t);
                } else {
                    List<String> strings = new ArrayList<String>();
                    strings.add(t);
                    map.put(key, strings);
                }
            }

            columnIndex++;
        }

        HSSFRow row;
        int rowNb = 1;
        for (String key : map.keySet()) {

            // Get or create row
            row = spreadSheet.createRow(rowNb);

            // Create cell keyword
            cell = row.createCell(0);
            cell.setCellValue(key);
            cell.setCellStyle(cellStyle);

            // Create cell values
            int columnNb = 1;
            for (String value : map.get(key)) {
                cell = row.createCell(columnNb);
                cell.setCellValue(new HSSFRichTextString(value));
                cell.setCellStyle(cellStyle);

                columnNb++;
            }

            rowNb++;
        }

        // Outputting to Excel spreadsheet
        FileOutputStream output;
        try {
            output = new FileOutputStream(new File(resEntry.getOutputFileName()) + ".xls");

            workbook.write(output);
            output.flush();
            output.close();

        } catch (FileNotFoundException e) {
            logger.log(Level.SEVERE, "File exists but is a directory rather than a regular file, does not exist but cannot be created, " +
                    "or cannot be opened for any other reason", e);
        } catch (IOException e) {
            logger.log(Level.SEVERE, "I/O error occurs", e);
        }
    }

    private String parseValue(String value) {
        value = value.replace("\\'", "'");
        value = value.replace("\\\"", "\"");
        value = value.replace("\\n", "\\\n");
        value = value.replace("\\", "");
        return value;
    }
}
