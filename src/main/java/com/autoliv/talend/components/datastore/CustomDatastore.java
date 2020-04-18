package com.autoliv.talend.components.datastore;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import org.talend.sdk.component.api.configuration.Option;
import org.talend.sdk.component.api.configuration.type.DataStore;
import org.talend.sdk.component.api.configuration.ui.layout.GridLayout;
import org.talend.sdk.component.api.meta.Documentation;

@DataStore("CustomDatastore")
@GridLayout({
    // the generated layout put one configuration entry per line,
    // customize it as much as needed
})
@Documentation("TODO fill the documentation for this configuration")
public class CustomDatastore implements Serializable {
    public static class ColumnFormats{
        @Option
        @Documentation("")
        public String Name;

        @Option
        @Documentation("")
        public int Ordered;
        public ColumnFormats setValue(String Name, int Ordered){
            this.Name = Name;
            this.Ordered = Ordered;
            return this;
        }
    }
    public static class totalColumn{
        @Option
        @Documentation("Column Name")
        public String ColName;

        @Option
        @Documentation("Column Prefix")
        public String ColPrefix;

        @Option
        @Documentation("Column format")
        public int ColFormat;

        public totalColumn setValue(String Name, String Prefix,int format){
            this.ColName = Name;
            this.ColPrefix = Prefix;
            this.ColFormat = format;
            return this;
        }
    }
    public static List<String> localSchema =  new ArrayList<>();


}