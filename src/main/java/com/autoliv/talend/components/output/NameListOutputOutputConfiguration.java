package com.autoliv.talend.components.output;

import java.io.Serializable;
import java.util.List;

import com.autoliv.talend.components.dataset.CustomDataset;

import com.autoliv.talend.components.datastore.CustomDatastore;
import org.talend.sdk.component.api.configuration.Option;
import org.talend.sdk.component.api.configuration.ui.layout.GridLayout;
import org.talend.sdk.component.api.configuration.ui.widget.Structure;
import org.talend.sdk.component.api.meta.Documentation;

@GridLayout({
    // the generated layout put one configuration entry per line,
    // customize it as much as needed
    //@GridLayout.Row({ "dataset" }),
    @GridLayout.Row({ "FileName" }),
    @GridLayout.Row({ "SheetName" }),
    @GridLayout.Row({ "config" }),
    @GridLayout.Row({ "ColumnFormat" }),
    @GridLayout.Row({ "AutoSizeColumn" })
})
@Documentation("TODO fill the documentation for this configuration")
public class NameListOutputOutputConfiguration implements Serializable {
    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private CustomDataset dataset;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String FileName;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String SheetName;

    @Option
    @Structure
    @Documentation("TODO fill the documentation for this parameter")
    private List<String> config;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private List<CustomDatastore.ColumnFormats> ColumnFormat;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private boolean AutoSizeColumn;

    public CustomDataset getDataset() {
        return dataset;
    }

    public NameListOutputOutputConfiguration setDataset(CustomDataset dataset) {
        this.dataset = dataset;
        return this;
    }

    public String getFileName() {
        return FileName;
    }

    public NameListOutputOutputConfiguration setFileName(String FileName) {
        this.FileName = FileName;
        return this;
    }

    public String getSheetName() {
        return SheetName;
    }

    public NameListOutputOutputConfiguration setSheetName(String SheetName) {
        this.SheetName = SheetName;
        return this;
    }
    public List<String>  getConfig() {
        return config;
    }

/*
    public String getConfig() {
        return config;
    }

    public NameListOutputOutputConfiguration setConfig(String config) {
        this.config = config;
        return this;
    }
*/
    public List<CustomDatastore.ColumnFormats> getColumnFormat() {

        return ColumnFormat;
    }

    public NameListOutputOutputConfiguration setColumnFormat(List<CustomDatastore.ColumnFormats> ColumnFormat) {
        this.ColumnFormat = ColumnFormat;
        return this;
    }

    public boolean getAutoSizeColumn() {
        return AutoSizeColumn;
    }

    public NameListOutputOutputConfiguration setAutoSizeColumn(boolean AutoSizeColumn) {
        this.AutoSizeColumn = AutoSizeColumn;
        return this;
    }
    public void updateSchemaToList(List<String> schemaList){
        if(schemaList.isEmpty()){
            return;
        }
        int colIndex = 0;
        for(String colName: schemaList){
            ColumnFormat.add((new CustomDatastore.ColumnFormats()).setValue(colName,colIndex++));
        }
    }

}