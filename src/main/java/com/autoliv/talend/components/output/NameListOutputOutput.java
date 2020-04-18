package com.autoliv.talend.components.output;

import static org.talend.sdk.component.api.component.Icon.IconType.CUSTOM;

import java.io.Serializable;

import javax.annotation.PostConstruct;
import javax.annotation.PreDestroy;

import com.bms.utils.ExcelTools;
import org.talend.sdk.component.api.component.Icon;
import org.talend.sdk.component.api.component.Version;
import org.talend.sdk.component.api.configuration.Option;
import org.talend.sdk.component.api.meta.Documentation;
import org.talend.sdk.component.api.processor.AfterGroup;
import org.talend.sdk.component.api.processor.BeforeGroup;
import org.talend.sdk.component.api.processor.ElementListener;
import org.talend.sdk.component.api.processor.Input;
import org.talend.sdk.component.api.processor.Processor;
import org.talend.sdk.component.api.record.Record;

import com.autoliv.talend.components.service.AutolivEtlService;

@Version(1) // default version is 1, if some configuration changes happen between 2 versions you can add a migrationHandler
@Icon(value = CUSTOM, custom = "NameListOutput") // icon is located at src/main/resources/icons/NameListOutput.svg
@Processor(name = "NameListOutput")
@Documentation("TODO fill the documentation for this processor")
public class NameListOutputOutput implements Serializable {
    private final NameListOutputOutputConfiguration configuration;
    private final AutolivEtlService service;
    private final ExcelTools exceltools;

    public NameListOutputOutput(@Option("configuration") final NameListOutputOutputConfiguration configuration,
                          final AutolivEtlService service) {
        this.configuration = configuration;
        this.service = service;
        exceltools = new ExcelTools(configuration.getFileName(),configuration.getSheetName());
    }

    @PostConstruct
    public void init() {
        // this method will be executed once for the whole component execution,
        // this is where you can establish a connection for instance
        // Note: if you don't need it you can delete it
    }

    @BeforeGroup
    public void beforeGroup() {
        // if the environment supports chunking this method is called at the beginning if a chunk
        // it can be used to start a local transaction specific to the backend you use
        // Note: if you don't need it you can delete it
    }

    @ElementListener
    public void onNext(
            @Input final Record defaultInput) {
        // this is the method allowing you to handle the input(s) and emit the output(s)
        // after some custom logic you put here, to send a value to next element you can use an
        // output parameter and call emit(value).
        exceltools.createSheet();

        exceltools.getDataFromRecord(defaultInput);
    }

    @AfterGroup
    public void afterGroup() {
        // symmetric method of the beforeGroup() executed after the chunk processing
        // Note: if you don't need it you can delete it
    }

    @PreDestroy
    public void release() {
        // this is the symmetric method of the init() one,
        // release potential connections you created or data you cached
        // Note: if you don't need it you can delete it
        try {

            //exceltools.printHeader(-1);

            if(configuration.getColumnFormat() != null ) {
                if (!configuration.getColumnFormat().isEmpty()) {
                    for (NameListOutputOutputConfiguration.ColumnFormats object : configuration.getColumnFormat()) {
                        exceltools.setColumnFormat(object.Name, object.Ordered);
                    }
                }
            }
            exceltools.printHeaderBySchema(this.configuration.getConfig(),-1);

            exceltools.printDatarowBySchema(this.configuration.getConfig(),-1);
            if(this.configuration.getAutoSizeColumn()) {
                exceltools.writeExcel(configuration.getFileName(),true);
                exceltools.reloadFile();
                exceltools.setAutoSizeCol();
            }
            //exceltools.printDatarow(-1);
            exceltools.writeExcel(configuration.getFileName(),true);
            exceltools.clearData();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}