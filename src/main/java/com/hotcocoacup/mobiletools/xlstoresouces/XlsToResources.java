package com.hotcocoacup.mobiletools.xlstoresouces;

import com.google.gson.Gson;
import com.google.gson.JsonIOException;
import com.google.gson.JsonSyntaxException;
import com.google.gson.reflect.TypeToken;
import com.hotcocoacup.mobiletools.xlstoresouces.model.Entry;
import com.hotcocoacup.mobiletools.xlstoresouces.model.KeyValuePair;
import org.apache.commons.cli.CommandLine;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

public class XlsToResources {

	public static final String VERSION = XlsToResources.class.getPackage().getImplementationVersion();
	public static final String LOGGER_NAME = "XlsToResources";

	private static Logger logger = Logger.getLogger(LOGGER_NAME);

	public void parse(String configFileName, CommandLine cmd) {

		logger.info("Reading configuration file " + configFileName);
		File file = new File(configFileName);
		Gson gson = new Gson();

		List<Entry> entries;
		try {
			entries = gson.fromJson(new FileReader(file),
					new TypeToken<List<Entry>>() {
					}.getType());
		} catch (JsonIOException e) {
			logger.log(Level.SEVERE, "Cannot parse the configuration file", e);
			return;
		} catch (JsonSyntaxException e) {
			logger.log(Level.SEVERE, "Cannot parse the configuration file", e);
			return;
		} catch (FileNotFoundException e) {
			logger.log(Level.SEVERE, "The configuration file does not exist", e);
			return;
		}

		logger.log(Level.INFO, entries.size()
				+ " entry(ies) found in the configuration file.");

		Map<String, List<KeyValuePair>> map = new HashMap<String, List<KeyValuePair>>();

		int entryCount = 1;
		for (Entry entry : entries) {
			Workbook workbook;

			logger.log(Level.INFO, "Entry #" + entryCount + ": Reading "
					+ entry.getXlsFile() + " ...");

			// parsing the excel file.
			try {
				if (entry.getXlsFile() == null) {
					logger.log(Level.SEVERE, "You must specify an XLS/XLSX file name. Ignoring the entry.");
					continue;
				}
				
				workbook = WorkbookFactory.create(new File(entry.getXlsFile()));
			} catch (InvalidFormatException e) {
				logger.log(Level.SEVERE,
						"Invalid file format. Ignoring this entry.", e);
				continue;
			} catch (IOException e) {
				logger.log(Level.SEVERE,
						"IO error while reading the file. Ignoring the entry.",
						e);
				continue;
			}
			
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

			// invalid sheet number
			if (entry.getSheet() < 0
					|| entry.getSheet() > workbook.getNumberOfSheets()) {
				logger.log(Level.SEVERE,
						"Sheet index not valid. Ignoring this entry.");
				continue;
			}

			Sheet sheet = workbook.getSheetAt(entry.getSheet());

			int rowEnd;
			if (entry.getRowEnd() == -1) {

				// default rowEnd : read all the rows
				rowEnd = sheet.getLastRowNum();
			} else {

				if (entry.getRowEnd() < 0
						|| entry.getRowEnd() < entry.getRowStart()) {
					logger.log(Level.SEVERE,
							"Invalid row end. Ignoring this entry.");
					continue;
				} else {
					rowEnd = Math.min(sheet.getLastRowNum(),
							entry.getRowEnd() - 1);
				}
			}

			// processing all the rows of the file
			for (int i = entry.getRowStart() - 1; i <= rowEnd; i++) {

				Row row = sheet.getRow(i);

				logger.log(Level.FINEST, " processing row: " + i + "...");
				
				if (row == null) {
					logger.log(Level.WARNING, " row: " + i + " is null");
					continue;
				}
				
				Cell keyCell = row.getCell(new CellReference(entry.getColumnKey()).getCol());
				Cell valueCell = row.getCell(new CellReference(entry.getColumnValue()).getCol());

				
				
				String keyStr = getString(evaluator, keyCell);
				String valueStr = getString(evaluator, valueCell);
				
				if (keyStr == null || keyStr.isEmpty()) {
					logger.log(Level.WARNING,
							"Key column " + entry.getColumnKey() + " (row "
									+ (i + 1)
									+ ") does not exist. Skipping row.");
					continue;
				}

				if (valueStr == null || valueStr.isEmpty()) {
					logger.log(Level.WARNING,
							"Value colum " + entry.getColumnValue() + " (row "
									+ (i + 1)
									+ ") does not exist. Skipping row.");
					continue;
				}

				String groupBy = "";
				if (entry.getGroupBy() != null) {
					Cell groupByCell = row.getCell(new CellReference(entry.getGroupBy()).getCol());

					if (groupByCell != null) {
						groupBy = groupByCell.getStringCellValue();
					} else {
						logger.log(
								Level.WARNING,
								"GroupBy column "
										+ entry.getGroupBy()
										+ " (row "
										+ (i + 1)
										+ ") does not exist. GroupBy set to default.");
					}
				}

				KeyValuePair keyValue = new KeyValuePair(keyStr, valueStr);

				add(map, groupBy, keyValue);
			}

			logger.log(Level.INFO, "Entry #" + entryCount
					+ ": Parsed with success.");

			entryCount++;
		}

		if (cmd.hasOption('a')) {
			
			String androidFileName = cmd.getOptionValue('a');
			logger.log(Level.INFO, "Exporting as android resource: " + androidFileName);
			
			Writer outputAndroidStream;
			try {
				outputAndroidStream = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(androidFileName), "UTF8"));
				Processor processorAndroid = new AndroidProcessor();
				processorAndroid.process(outputAndroidStream, map);
				logger.log(Level.INFO, "Exported with success");
			} catch (IOException e) {
				logger.log(Level.SEVERE, "Export failed...", e);
			}

		}
		
		if (cmd.hasOption('i')) {
			String iosFileName = cmd.getOptionValue('i');
			logger.log(Level.INFO, "Exporting as ios resource: " + iosFileName);
			
			Writer outputIosStream;
			try {
				outputIosStream = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(iosFileName), "UTF8"));
				Processor processorIos = new IosProcessor();
				processorIos.process(outputIosStream, map);
				logger.log(Level.INFO, "Exported with success");
			} catch (IOException e) {
				logger.log(Level.SEVERE, "Export failed...", e);
			}
			
		}
		
		logger.log(Level.INFO, "End of execution");
	}
	
	private static String getString(FormulaEvaluator evaluator, Cell cell) {
		
		if (cell == null) {
			return "";
		}
		
		CellValue cellValue = evaluator.evaluate(cell);
		
		if (cellValue == null) {
			return "";
		}
		
		switch (cellValue.getCellType()) {
		
			case Cell.CELL_TYPE_BOOLEAN:
				return cellValue.getBooleanValue() ? "true" : "false";
			case Cell.CELL_TYPE_NUMERIC:
				return String.valueOf(cellValue.getNumberValue());
			case Cell.CELL_TYPE_STRING:
				return cellValue.getStringValue();
			case Cell.CELL_TYPE_FORMULA: // not happening because we evaluate
			case Cell.CELL_TYPE_ERROR:
			case Cell.CELL_TYPE_BLANK:
			default:
				return "";
		}
		
	}

	private static void add(Map<String, List<KeyValuePair>> map,
			String groupBy, KeyValuePair keyValue) {

		List<KeyValuePair> list;
		if (!map.containsKey(groupBy)) {
			list = new ArrayList<KeyValuePair>();
			map.put(groupBy, list);
		} else {
			list = map.get(groupBy);
		}

		list.add(keyValue);
	}

}
