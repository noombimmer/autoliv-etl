package com.autoliv.talend.components.output;

import java.io.Serializable;

import com.autoliv.talend.components.dataset.CustomDataset;

import org.talend.sdk.component.api.configuration.Option;
import org.talend.sdk.component.api.configuration.ui.layout.GridLayout;
import org.talend.sdk.component.api.meta.Documentation;

@GridLayout({
    // the generated layout put one configuration entry per line,
    // customize it as much as needed
    @GridLayout.Row({ "dataset" }),
    @GridLayout.Row({ "FileName" }),
    @GridLayout.Row({ "SheetName" })
})
@Documentation("TODO fill the documentation for this configuration")
public class HRHCSummaryOutputOutputConfiguration implements Serializable {
    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private CustomDataset dataset;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String FileName;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String SheetName;

    public CustomDataset getDataset() {
        return dataset;
    }

    public HRHCSummaryOutputOutputConfiguration setDataset(CustomDataset dataset) {
        this.dataset = dataset;
        return this;
    }

    public String getFileName() {
        return FileName;
    }

    public HRHCSummaryOutputOutputConfiguration setFileName(String FileName) {
        this.FileName = FileName;
        return this;
    }

    public String getSheetName() {
        return SheetName;
    }

    public HRHCSummaryOutputOutputConfiguration setSheetName(String SheetName) {
        this.SheetName = SheetName;
        return this;
    }
}