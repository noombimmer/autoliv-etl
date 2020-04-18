package com.autoliv.talend.components.processor;

import java.io.Serializable;

import org.talend.sdk.component.api.configuration.Option;
import org.talend.sdk.component.api.configuration.ui.layout.GridLayout;
import org.talend.sdk.component.api.meta.Documentation;

@GridLayout({
    // the generated layout put one configuration entry per line,
    // customize it as much as needed
    @GridLayout.Row({ "schemaList" }),
    @GridLayout.Row({ "ColumnList" })
})
@Documentation("TODO fill the documentation for this configuration")
public class AutolivProcessorProcessorConfiguration implements Serializable {
    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String schemaList;

    @Option
    @Documentation("TODO fill the documentation for this parameter")
    private String ColumnList;

    public String getSchemaList() {
        return schemaList;
    }

    public AutolivProcessorProcessorConfiguration setSchemaList(String schemaList) {
        this.schemaList = schemaList;
        return this;
    }

    public String getColumnList() {
        return ColumnList;
    }

    public AutolivProcessorProcessorConfiguration setColumnList(String ColumnList) {
        this.ColumnList = ColumnList;
        return this;
    }
}