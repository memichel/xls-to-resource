package com.hotcocoacup.mobiletools.xlstoresouces;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

import java.util.logging.ConsoleHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created by a556679 on 17/02/2016.
 */
public class ConverterMain {

    public static final String VERSION = ConverterMain.class.getPackage().getImplementationVersion();
    public static final String LOGGER_NAME = "ConverterMain";

    private static Logger logger = Logger.getLogger(LOGGER_NAME);
    private static Options options = new Options();

    public static void main(String[] args) {

        // Setting up the logger
        logger.setLevel(Level.INFO);
        logger.setUseParentHandlers(false);

        ConsoleHandler handler = new ConsoleHandler();
        handler.setFormatter(new LogFormatter());
        logger.addHandler(handler);

        // Parsing the user inputs
        options.addOption("h", "help", false, "Print the help.");
        options.addOption("v", "version", false, "Print the current version.");
        options.addOption("c", "config", true, "The configuration file");
        options.addOption("a", "android", true, "The android resouce filename to export");
        options.addOption("i", "ios", true, "The iOS resource filename to export");
        options.addOption("e", "export", true, "The conversion target ('rtx' for resources to xls OR 'xtr' for xls to resources)");

        CommandLineParser parser = new BasicParser();
        CommandLine cmd;

        try {
            cmd = parser.parse(options, args);
        } catch (ParseException e) {
            logger.log(Level.SEVERE, "Failed to parse command line properties",
                    e);
            help();
            return;
        }

        // user asked for help...
        if (cmd.hasOption('h')) {
            help();
            return;
        }

        // user asked for version
        if (cmd.hasOption('v')) {
            printVersion();
            return;
        }

        // extracting the configuration filename
        String configFileName;
        if (cmd.hasOption('c')) {
            configFileName = cmd.getOptionValue('c');
        } else {
            logger.severe("You must input the configurationFilename");
            help();
            return;
        }

        if (cmd.hasOption('e')) {
            if (cmd.getOptionValue('c').equals("xtr")) {
                XlsToResources xlsToResources = new XlsToResources();
                xlsToResources.parse(configFileName, cmd);
            } else {
                ResourcesToXls excel = new ResourcesToXls();
                excel.generateExcel(configFileName);
            }
        }
    }

    private static void helpExtractResource() {
        HelpFormatter formater = new HelpFormatter();
        formater.printHelp("Main", options);

        System.out.println("\nFormat of the Configuration file:");
        System.out.println("[");
        System.out.println("     {");
        System.out.println("         \"outputFileName\": \"xls file name. default : Wording.xls\",");
        System.out.println("         \"sheetName\": \"sheet name, default : Wording\",");
        System.out.println("         \"firstColumnName\": \"first column name, default : Reference\",");
        System.out.println("         \"resourcesFiles\": [");
        System.out.println("                {");
        System.out.println("                    \"fileName\": \"resource file containing the wording. Mandatory\",");
        System.out.println("                    \"columnName\": \"column name for this file\"");
        System.out.println("                }, ...");
        System.out.println("         ]");
        System.out.println("     }, ...");
        System.out.println("]");
        System.out.println("");
        System.out.println("Example of how to use:");
        System.out.println("java -jar xlsToResource.jar -c config-res-to-xls.json -a string.xml -i sample.strings -e xtr");

        System.exit(0);
    }

    private static void help() {
        HelpFormatter formater = new HelpFormatter();
        formater.printHelp("Main", options);

        System.out.println("\nFormat of the Configuration file:");
        System.out.println("[");
        System.out.println("     {");
        System.out.println("         \"fileName\": (string) \"xls or xlsx file containing the wording. Mandatory.\",");
        System.out.println("         \"sheet\": (int) \"index of the sheet concerned. 0=first sheet. Default=0\", ");
        System.out.println("         \"rowStart\": (int) \"index of the starting row. 1=first row. Default=1\", ");
        System.out.println("         \"rowEnd\": (int) \"index of the last row. 1=first row. -1=all rows. Default=-1\", ");
        System.out.println("         \"columnKey\": (String) \"letter of the column containing the key. . Default='A'\", ");
        System.out.println("         \"columnValue\": (String) \"letter of the column containing the value. Default='B'\", ");
        System.out.println("         \"groupBy\": (String) \"letter of the column containing the group value. null=Do not group. Default=null\", ");
        System.out.println("     }, ...");
        System.out.println("]");
        System.out.println("");
        System.out.println("Example of how to use:");
        System.out.println("java -jar xlsToResource.jar -c config-xls-to-res.json -a string.xml -i sample.strings -e xtr");

        System.exit(0);
    }

    private static void printVersion() {
        System.out.println("V" + VERSION);
    }
}
