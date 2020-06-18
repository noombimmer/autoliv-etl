package com.autoliv.talend.components.output;

import static org.talend.sdk.component.api.component.Icon.IconType.CUSTOM;

import java.io.Serializable;
import java.util.Map;

import javax.annotation.PostConstruct;
import javax.annotation.PreDestroy;

import com.autoliv.talend.components.datastore.CustomDatastore;
import com.bms.utils.ExcelTools;
import com.bms.utils.PivotTools;
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
@Icon(value = CUSTOM, custom = "UploadHRMOutput") // icon is located at src/main/resources/icons/UploadHRMOutput.svg
@Processor(name = "UploadHRMOutput")
@Documentation("TODO fill the documentation for this processor")
public class UploadHRMOutputOutput implements Serializable {
    private final UploadHRMOutputOutputConfiguration configuration;
    private final AutolivEtlService service;
    private final PivotTools pivottools;
    public UploadHRMOutputOutput(@Option("configuration") final UploadHRMOutputOutputConfiguration configuration,
                          final AutolivEtlService service) {
        this.configuration = configuration;
        this.service = service;
        pivottools = new PivotTools(configuration.getFileName(),configuration.getSheetName());
    }

    @PostConstruct
    public void init() {
        // this method will be executed once for the whole component execution,
        // this is where you can establish a connection for instance
        // Note: if you don't need it you can delete it
        pivottools.setLocalSchemaList(this.configuration.getConfig());
        if(configuration.getGrandTotalColumn()){
            pivottools.setGrandTotal = true;
        }

        if(configuration.getGroupTotalColumn() != null){
            if(!configuration.getGroupTotalColumn().isEmpty()){
                for(CustomDatastore.totalColumn item:configuration.getGroupTotalColumn()){
                    if(item.ColFormat == 0){
                        pivottools.GroupTotalCol.put("GroupCut",item.ColName);
                        pivottools.GroupTotalCol.put("GroupSuffix",item.ColPrefix);
                        pivottools.activeGroupTotal = true;
                    }else if(item.ColFormat == 1){
                        pivottools.columnRename.put(item.ColName,item.ColPrefix);
                        pivottools.activeRenameColumn = true;
                    }else if(item.ColFormat == 2){
                        pivottools.GroupTotalCol.put("GroupCodeTile",item.ColName);
                    }else if(item.ColFormat == 3){
                        pivottools.GroupTotalCol.put("GroupCodeDescription",item.ColName);
                    } else if(item.ColFormat == 4){
                        pivottools.GroupTotalCol.put("GroupCodeTilePosition",item.ColName);
                    }else if(item.ColFormat == 5){
                        pivottools.GroupTotalCol.put("GrantotalCol",item.ColName);
                        pivottools.GroupTotalCol.put("GrantotalTitle",item.ColPrefix);
                    }

                }
            }
        }
        if(configuration.getColumnFormat() != null ) {
            if (!configuration.getColumnFormat().isEmpty()) {
                for (CustomDatastore.ColumnFormats object : configuration.getColumnFormat()) {
                    pivottools.setColumnFormat(object.Name, object.Ordered);
                }
            }
        }
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
        pivottools.createSheet();

        pivottools.getDataFromRecord(defaultInput);

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
            if(pivottools.GroupTotalCol.get("GroupCut") != null){
                pivottools.groupTotalLast();
                pivottools.updateGroupTotal();
            }
            pivottools.printHeaderBySchema(this.configuration.getConfig(),-1);
            pivottools.printDatarowBySchema(this.configuration.getConfig(),-1);
            pivottools.printRow();
            if(this.configuration.getAutoSizeColumn()) {
                pivottools.writeExcel(configuration.getFileName(),true);
                pivottools.reloadFile();
                pivottools.setAutoSizeCol();
            }
            //exceltools.printDatarow(-1);
            pivottools.writeExcel(configuration.getFileName(),true);
            pivottools.clearData();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}