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
    @GridLayout.Row({ "AutoSizeColumn" }),
    @GridLayout.Row({ "GroupTotalEmpColumn" }),
    @GridLayout.Row({ "GroupTotalTempColumn" }),
    @GridLayout.Row({ "GroupTotalColumn" }),
    @GridLayout.Row({ "GrandTotalEmpColumn" }),
    @GridLayout.Row({ "GrandTotalTempColumn" }),
    @GridLayout.Row({ "GrandTotalColumn" })
})
@Documentation("TODO fill the documentation for this configuration")
public class UploadHRMOutputOutputConfiguration implements Serializable {
    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private CustomDataset dataset;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String FileName;

    @Option
    @Documentation("Sheet Name ")
    private String SheetName;

    @Option
    @Structure
    @Documentation("TODO fill the documentation for this parameter")
    private List<String> config;

    @Option
    @Documentation("ColumnFormat ")
    private List<CustomDatastore.ColumnFormats> ColumnFormat;

    @Option
    @Documentation("AutoSizeColumn True/False")
    private boolean AutoSizeColumn;

    @Option
    @Documentation("GrandTotalColumn True/False")
    private boolean GrandTotalColumn;


    @Option
    @Documentation("ColumnFormat ")
    private List<CustomDatastore.totalColumn> GroupTotalEmpColumn;

    @Option
    @Documentation("ColumnFormat ")
    private List<CustomDatastore.totalColumn> GroupTotalTempColumn;

    @Option
    @Documentation("ColumnFormat ")
    private List<CustomDatastore.totalColumn> GroupTotalColumn;

    @Option
    @Documentation("ColumnFormat ")
    private List<CustomDatastore.totalColumn> GrandTotalEmpColumn;

    @Option
    @Documentation("ColumnFormat ")
    private List<CustomDatastore.totalColumn> GrandTotalTempColumn;


    public boolean getGrandTotalColumn() {
        return GrandTotalColumn;
    }


    public List<String>  getConfig() {
        return config;
    }

    public List<CustomDatastore.totalColumn> getGroupTotalColumn(){
        return GroupTotalColumn;
    };
    public List<CustomDatastore.totalColumn> getGroupTotalEmpColumn(){
        return GroupTotalEmpColumn;
    };
    public List<CustomDatastore.totalColumn> getGroupTotalTempColumn(){
        return GroupTotalTempColumn;
    };
    public List<CustomDatastore.totalColumn> getGrandTotalEmpColumn(){
        return GrandTotalEmpColumn;
    };
    public List<CustomDatastore.totalColumn> getGrandTotalTempColumn(){
        return GrandTotalTempColumn;
    };

    public CustomDataset getDataset() {
        return dataset;
    }

    public UploadHRMOutputOutputConfiguration setDataset(CustomDataset dataset) {
        this.dataset = dataset;
        return this;
    }

    public String getFilrName() {
        return FileName;
    }

    public UploadHRMOutputOutputConfiguration setFilrName(String FilrName) {
        this.FileName = FilrName;
        return this;
    }

    public String getSheetName() {
        return SheetName;
    }

    public UploadHRMOutputOutputConfiguration setSheetName(String SheetName) {
        this.SheetName = SheetName;
        return this;
    }
    public String getFileName() {
        return FileName;
    }

    public UploadHRMOutputOutputConfiguration setFileName(String FileName) {
        this.FileName = FileName;
        return this;
    }
    public boolean getAutoSizeColumn() {
        return AutoSizeColumn;
    }

    public UploadHRMOutputOutputConfiguration setAutoSizeColumn(boolean AutoSizeColumn) {
        this.AutoSizeColumn = AutoSizeColumn;
        return this;
    }
    public List<CustomDatastore.ColumnFormats> getColumnFormat() {

        return ColumnFormat;
    }

    public UploadHRMOutputOutputConfiguration setColumnFormat(List<CustomDatastore.ColumnFormats> ColumnFormat) {
        this.ColumnFormat = ColumnFormat;
        return this;
    }
}