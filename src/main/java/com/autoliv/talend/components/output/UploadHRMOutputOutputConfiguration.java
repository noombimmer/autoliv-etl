package com.autoliv.talend.components.output;

import java.io.Serializable;

import com.autoliv.talend.components.dataset.CustomDataset;

import org.talend.sdk.component.api.configuration.Option;
import org.talend.sdk.component.api.configuration.ui.layout.GridLayout;
import org.talend.sdk.component.api.meta.Documentation;

@GridLayout({
    // the generated layout put one configuration entry per line,
    // customize it as much as needed
    //@GridLayout.Row({ "dataset" }),
    @GridLayout.Row({ "FilrName" }),
    @GridLayout.Row({ "SheetName" }),
    @GridLayout.Row({ "config" }),
    @GridLayout.Row({ "ColumnFormat" }),
    @GridLayout.Row({ "AutoSizeColumn" }),
    @GridLayout.Row({ "GroupTotalColumn" }),
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
    private String FilrName;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String SheetName;

    public CustomDataset getDataset() {
        return dataset;
    }

    public UploadHRMOutputOutputConfiguration setDataset(CustomDataset dataset) {
        this.dataset = dataset;
        return this;
    }

    public String getFilrName() {
        return FilrName;
    }

    public UploadHRMOutputOutputConfiguration setFilrName(String FilrName) {
        this.FilrName = FilrName;
        return this;
    }

    public String getSheetName() {
        return SheetName;
    }

    public UploadHRMOutputOutputConfiguration setSheetName(String SheetName) {
        this.SheetName = SheetName;
        return this;
    }
}