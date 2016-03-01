package com.hotcocoacup.mobiletools.xlstoresouces.model;

import java.util.List;

/**
 * Created by a556679 on 25/02/2016.
 */
public class ResEntry {
    private String outputFileName;
    private String sheetName;
    private String firstColumnName;
    private List<ResFileEntry> resourcesFiles;

    /**
     * Get the outputFileName field
     *
     * @return outputFileName
     */
    public String getOutputFileName() {
        return outputFileName;
    }

    /**
     * Get the sheetName field
     *
     * @return sheetName
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * Get the firstColumnName field
     *
     * @return firstColumnName
     */
    public String getFirstColumnName() {
        return firstColumnName;
    }

    /**
     * Get the resourcesFiles field
     *
     * @return resourcesFiles
     */
    public List<ResFileEntry> getResourcesFiles() {
        return resourcesFiles;
    }
}
