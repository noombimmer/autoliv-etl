package com.bms.utils;


import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.talend.sdk.component.api.record.Record;
import org.talend.sdk.component.api.record.Schema;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PivotToolsHUCSUM {
    private static File fileInstanced = null;
    private static FileInputStream fIP = null;
    public boolean fixrowHeader;
    public int NumHeaderRow;
    public boolean rowSummary;
    public boolean GroupHeader;
    private String lastCut;
    private String lFileName;
    private String lSheetName = "Sheet1";
    private static Workbook wb = null;
    private static Sheet sheet = null;
    private int rowAccessWindowSize = SXSSFWorkbook.DEFAULT_WINDOW_SIZE;// used in auto flush
    private boolean appendWorkbook = false;
    private boolean appendSheet = false;
    private boolean recalculateFormula = false;
    public boolean activeGroupTotal = false;
    public boolean activeRenameColumn = false;

    public boolean setHeaderColor = true;
    public boolean setHeaderBorder = true;
    public boolean setHeaderBold = true;
    public boolean setHeaderRowHeight = true;
    public boolean setHeaderTextCenter = true;
    public boolean setGrandTotal = false;
    private boolean isRequireFormatStyle = false;
    private boolean isRequireFormatGroupTotal = false;
    public static int RowTotalColStart;
    private double RowTotalBuffer;
    private String empTypeVal;
/*    private boolean isRequireFormatStyle_Emp;
    private boolean isRequireFormatStyle_Temp;
    private boolean isRequireFormatGroupTotal_Emp;
    private boolean isRequireFormatGroupTotal_Temp;*/
    private XSSFColor colorCodeGrandTotalEmpStyle = null;
    private XSSFColor colorCodeGrandTotalTempStyle = null;
    private XSSFColor colorCodeGroupTotalTempStyle = null;
    private XSSFColor colorCodeGroupTotalEmpStyle = null;


    private int colCount = 0;
    private int rowCount = 0;
    private int intHeaderRow = 0;
    private int intDataRow = 0;
    private int globalColIndex = 0;

    private Record dataRows;
    private Row row;
    private Map<String, Object> globalColumns = new HashMap<String, Object>();
    private Map<String, Object> formatColumns = new HashMap<String, Object>();
    private Map<String, rowData> rsRows = new HashMap<String, rowData>();

    private Map<String, Object> GrandTotalVal = new HashMap<String, Object>();
    private Map<String, Object> GrandTotalEmpVal = new HashMap<String, Object>();
    private Map<String, Object> GrandTotalTempVal = new HashMap<String, Object>();

    private Map<String, Double> GroupTotalVal = new HashMap<String, Double>();
    private Map<String, Double> GroupTotalEmpVal = new HashMap<String, Double>();
    private Map<String, Double> GroupTotalTempVal = new HashMap<String, Double>();

    public static Map<String, Object> GroupTotalCol = new HashMap<String, Object>();
    public static Map<String, Object> GroupTotalCutVal = new HashMap<String, Object>();
/*    public static Map<String, Object> GroupTotalEmpCutVal = new HashMap<String, Object>();
    public static Map<String, Object> GroupTotalTempCutVal = new HashMap<String, Object>();*/
    public static Map<String, Object> columnRename = new HashMap<String, Object>();
    public static Map<String, Object> rowAppend = new HashMap<String, Object>();
/*    public static Map<String, Object> rowAppendEmp = new HashMap<String, Object>();
    public static Map<String, Object> rowAppendTemp = new HashMap<String, Object>();*/

    public static Map<String, String> HeaderGroup = new HashMap<String, String>();

    private Map<String, rowData> rsTotalRows = new HashMap<String, rowData>();

    private XSSFCellStyle lHeaderStyle = null;
    private XSSFCellStyle lDetailsStyle = null;
    public static int HeaderGroiupCount = 0;
/*
    private XSSFCellStyle lDetailsGrandTotalStyle = null;
    private XSSFCellStyle lDetailsGrandTotalEmpStyle = null;
    private XSSFCellStyle lDetailsGrandTotalTempStyle = null;
*/
/*

    private XSSFCellStyle lDetailsGroupTotalStyle = null;
    private XSSFCellStyle lDetailsGroupTotalEmpStyle = null;
    private XSSFCellStyle lDetailsGroupTotalTempStyle = null;
*/

    private XSSFCellStyle lDetailsStyleCenter = null;
    public List<String> localSchemaList;

    public short shortCurrencyFormat = 0;
    //public short shortQtyFormat = 0;
    public short shortSingleDigit = 0;

    private Font fnt = null;
    private Font fnt2 = null;
    private String lastGroupCode = null;
    private XSSFColor colorCodeGroupTotalStype = null;
    private XSSFColor colorCodeGrandTotalStyle = null;
    private XSSFColor colorCodeHeader = null;
    private String cutValue;
    private Object lDetailsStyleLeft;
    DecimalFormat decimalFormat = new DecimalFormat("#,##0");

    public void setColumnFormat(String colName, int formatNum) {

        formatColumns.put(colName, formatNum);
    }

    public void reloadFile() {
        if (this.lFileName.isEmpty()) return;
        if (wb == null) return;
        if (sheet == null) return;
        if (this.fIP != null) {
            try {
                this.fIP.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        if (fileInstanced.exists()) {
            try {
                this.fIP = new FileInputStream(this.fileInstanced);
                wb = new XSSFWorkbook(fIP);
                sheet = wb.getSheet(lSheetName);
                lDetailsStyle = null;
                lHeaderStyle = null;
                shortCurrencyFormat = 0;

            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            return;
        }
    }

    public PivotToolsHUCSUM(String fileName, String sheetName) {
        this.lFileName = fileName;
        ZipSecureFile.setMinInflateRatio(0);
        if (!sheetName.isEmpty() && sheetName.length() > 0 && sheetName != null) {
            this.lSheetName = sheetName;
        }
        fileInstanced = new File(this.lFileName);
        if (fileInstanced.exists()) {
            try {
                this.fIP = new FileInputStream(this.fileInstanced);
                wb = new XSSFWorkbook(fIP);
                sheet = wb.getSheet(lSheetName);
                if (sheet != null) {
                    wb.removeSheetAt(wb.getSheetIndex(lSheetName));
                    sheet = wb.createSheet(lSheetName);
                } else {
                    sheet = wb.createSheet(lSheetName);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            fileInstanced.delete();
            wb = new SXSSFWorkbook(rowAccessWindowSize);
            sheet = wb.createSheet(lSheetName);
        }
        //RowTotalColStart = 2;
    }

    public void writeExcel(String fileName, boolean createDir) throws Exception {
        System.out.print("\b\b\b\b Save File : " + fileName + "\r\n");
        if (createDir) {
            File file = new File(fileName);
            File pFile = file.getParentFile();
            if (pFile != null && !pFile.exists()) {
                pFile.mkdirs();
            }
        }
        FileOutputStream fileOutput = new FileOutputStream(fileName);
        if (appendWorkbook && appendSheet && recalculateFormula) {
            evaluateFormulaCell();
        }
        wb.write(fileOutput);
        fileOutput.close();
        lDetailsStyle = null;
        lHeaderStyle = null;
        shortCurrencyFormat = 0;

    }

    private void evaluateFormulaCell() {
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
            sheet = wb.getSheetAt(sheetNum);
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellTypeEnum() == CellType.FORMULA) {
                        evaluator.evaluateFormulaCellEnum(c);
                    }
                }
            }
        }
    }

    public void createSheet() {
        this.sheet = wb.getSheet(this.lSheetName);
        if (this.sheet == null) {
            wb.createSheet(this.lSheetName);
        }
    }

    public void setAutoSizeCol() {
        if (this.sheet != null) {
            System.out.print("\b\b\b\b Set AutoSize Columns\r\n");
            int colIndex = 0;
            for (String key : localSchemaList) {
                //HeaderGroup.put("fix_width_" + String.valueOf(colIndex),"1");
                if(HeaderGroup.containsKey("fix_width_" + String.valueOf(colIndex))){
                    System.out.println("fix_width_ :" + key + ", " + "fix_width_" + String.valueOf(colIndex ));
                    this.sheet.setColumnWidth(colIndex,2000);
                }else {
                    System.out.println("Set column :" + key + " Autosize " + String.valueOf(colIndex ));
                    this.sheet.autoSizeColumn(colIndex,true);
                }
                colIndex++;

            }
        }
    }

    public void clearData() {
        colCount = 0;
        rowCount = 0;
        dataRows = null;
        intHeaderRow = 0;
        intDataRow = 0;
        row = null;

        globalColumns.clear();
        formatColumns.clear();
        rsRows.clear();
        localSchemaList.clear();

        GrandTotalVal.clear();
        GrandTotalEmpVal.clear();
        GrandTotalTempVal.clear();

        GroupTotalVal.clear();
        GroupTotalEmpVal.clear();
        GroupTotalTempVal.clear();

        GroupTotalCol.clear();
        columnRename.clear();
        rsTotalRows.clear();

        globalColIndex = 0;
        lHeaderStyle = null;
        lDetailsStyle = null;

        File fileInstanced = null;
        FileInputStream fIP = null;
        lFileName = "";
        lSheetName = "Sheet1";
        wb = null;
        sheet = null;

        appendWorkbook = false;
        appendSheet = false;
        recalculateFormula = false;

        setHeaderColor = true;
        setHeaderBorder = true;
        setHeaderBold = true;
        setHeaderRowHeight = true;
        setHeaderTextCenter = true;

        lDetailsStyleCenter = null;

/*
        lDetailsGrandTotalStyle = null;
        lDetailsGrandTotalEmpStyle = null;
        lDetailsGrandTotalTempStyle = null;
*/

/*
        lDetailsGroupTotalStyle = null;
        lDetailsGroupTotalEmpStyle = null;
        lDetailsGroupTotalTempStyle = null;
*/

        activeRenameColumn = false;
        activeGroupTotal = false;
        fnt = null;
        fnt2 = null;
        lastCut = null;
        System.out.println("End Clear Data");
        lDetailsStyleQty = null;
        tempHeaderGroupName = "";
    }

    public void printRow() {
        //System.err.printf(" rowCount  : %d\n", rowCount);
        for (String tt : rowAppend.keySet()) {
            System.err.printf("rowAppend : %s : [%s ]\n", tt, rowAppend.get(tt).toString());
        }
        /*
        for (String KeyRow : rsRows.keySet()) {
            //System.err.printf("rsRows %s : [%s ]\n", KeyRow, rsRows.get(KeyRow).toString());
            if (rowAppend.containsKey(KeyRow)) {
                System.err.printf("%s rowAppend : [%s ]\n", KeyRow, (String) rowAppend.get(KeyRow));
                for (String key2 : rsTotalRows.get((String) rowAppend.get(KeyRow)).rsColumns.keySet()) {
                    System.err.printf("%s : [%s ]", key2, rsTotalRows.get((String) rowAppend.get(KeyRow)).rsColumns.get(key2).toString());
                }
                System.err.printf("\n");
            }
        }

         */

    }

    public void setColumnWidth() {
        sheet.setColumnWidth(3,1500);
        sheet.setColumnWidth(4,2000);
    }

    public class rowData {
        public Map<String, Object> rsColumns = new HashMap<String, Object>();

        public rowData() {
        }

        public void addColumn(String colName, Object colValue) {
            if (rsColumns.get(colName) == null) {
                rsColumns.put(colName, colValue);
            }
            addGlobalColumn(colName);
        }
    }

    public void addGlobalColumn(String colName) {

        if (globalColumns.get(colName) == null) {
            globalColumns.put(colName, globalColIndex++);
        }
    }

    public void getDataFromRecord(Record datainput) {
        colCount = 0;
        dataRows = datainput;
        List<Schema.Entry> arrayRec = datainput.getSchema().getEntries();

        rsRows.put(String.valueOf(rowCount), new rowData());

        arrayRec.forEach(this::workwithData);

        //arrayRec.get(0).getName();
        for (String keyCol : localSchemaList) {
            rowData data = rsRows.get(String.valueOf(rowCount));
            if (data.rsColumns.containsKey(keyCol)) {
                Object value = data.rsColumns.get(keyCol);
                Object cValue = value;
                if(GroupTotalCol.containsKey("GroupCodeTile")) {
                    if (((String) GroupTotalCol.get("GroupCodeTile")).equalsIgnoreCase(keyCol)) {
                        lastGroupCode = String.valueOf(cValue);
                    }
                    if (((String) GroupTotalCol.get("GroupCut")).equalsIgnoreCase(keyCol)) {
                        this.cutValue = cValue != null ? cValue.toString() : "";
                        if (lastCut != null) {
                             if (lastCut.equalsIgnoreCase(String.valueOf(cValue))) {

                            } else {
                                groupTotalLast();
                                updateGroupTotal();
                                lastCut = String.valueOf(cValue);
                            }
                        } else {
                            lastCut = cValue != null ? cValue.toString() : "";
                        }
                    }
                    GroupTotalAddValue(keyCol, aDoubleTryParse(value != null ? value.toString() : "0"), colCount);

                }
                if (setGrandTotal) {
                    appendGrandTotal(keyCol, aDoubleTryParse(String.valueOf(cValue)));
                }
            }

        }
        rowCount++;
    }

    public void printHeaderBySchema(final List<String> SchemaList, int atRow) {

        localSchemaList = SchemaList;
        firstheaderRow = null;
        int colIndex = 0;
        if (atRow >= 0) {
            intHeaderRow = atRow;
        }
        if(fixrowHeader == true){
            intHeaderRow = NumHeaderRow - 1;
        }
        if (lHeaderStyle == null) {
            lHeaderStyle = (XSSFCellStyle) wb.createCellStyle();

            if (fnt == null) {
                fnt = wb.createFont();
            }
            if (this.setHeaderBold) fnt.setBold(true);
            if (this.setHeaderBold && this.setHeaderColor) fnt.setColor(IndexedColors.WHITE.getIndex());
            lHeaderStyle.setFont(fnt);
            if (this.setHeaderColor) {
                if (colorCodeHeader == null) {
                    try {
                        colorCodeHeader = new XSSFColor((byte[]) Hex.decodeHex("0070c0"), null);
                    } catch (DecoderException e) {
                        e.printStackTrace();
                    }
                }
                lHeaderStyle.setFillBackgroundColor(colorCodeHeader);
                lHeaderStyle.setFillForegroundColor(colorCodeHeader);
                lHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }

        for (String key : SchemaList) {
            String strName = "";
            if (row == null) {
                row = this.sheet.createRow(intHeaderRow);
            }
            strName = key;
            if(NumHeaderRow > 1){

                if(!getHeaderGroupPrefix(strName)) {
                    int localColIndex = colIndex++;
                    if (activeRenameColumn) {
                        strName = this.columnRename.get(key) == null ? key : String.valueOf(this.columnRename.get(key));

                    }else{
                        strName = key;
                    }
                    cellMerge(strName, localColIndex, localColIndex, intHeaderRow - 1, intHeaderRow);
                }else{

                    if(!HeaderGroup.containsKey(tempHeaderGroupName + "_merged")){

                        int localColStart = Integer.parseInt(HeaderGroup.get(tempHeaderGroupName + "_colStart"));
                        int localColEnd = Integer.parseInt(HeaderGroup.get(tempHeaderGroupName + "_colEnd"));
                        cellMerge(tempHeaderGroupName, localColStart, localColEnd, intHeaderRow -1, intHeaderRow -1);
                        HeaderGroup.put(tempHeaderGroupName + "_merged","1");

                    }
                    strName = key;
                    HeaderGroup.put("fix_width_" + String.valueOf(colIndex),"1");
                    if (activeRenameColumn) {
                        strName = this.columnRename.get(key) == null ? key : String.valueOf(this.columnRename.get(key));
                        whiteHeader(strName, colIndex++);
                    }else {
                        whiteHeader(strName, colIndex++);
                    }

                }
            }else{
                if (activeRenameColumn) {
                    strName = this.columnRename.get(key) == null ? key : String.valueOf(this.columnRename.get(key));
                    whiteHeader(strName, colIndex++);
                } else {
                    strName = key;
                    whiteHeader(strName, colIndex++);
                }
            }
        }
        if(this.rowSummary == true) {
            whiteHeader("Total", colIndex++);
            if(NumHeaderRow > 1){
                int localColIndex = colIndex - 1;
                cellMerge("Total",localColIndex, localColIndex,intHeaderRow -1,intHeaderRow);
            }
        }
    }
    public Row firstheaderRow = null;
    public String tempHeaderGroupName = "";
    public boolean getHeaderGroupPrefix(String colName){
        boolean result = false;
        for(String key: HeaderGroup.keySet()){
            //"_prefix"
            if(key.matches("(.*)_prefix")) {
                if (colName.matches(HeaderGroup.get(key) + "(.*)") && result == false) {
                    tempHeaderGroupName = key.split("_")[0];
                    result = true;
                    break;
                }
            }
        }
        return result;
    }
    public void setHeaderGroupColumn(String colName,int colIndex){
        boolean result = false;
        for(String key: HeaderGroup.keySet()){
            //"_prefix"
            if(key.matches("(.*)_prefix")) {
                if (colName.matches(HeaderGroup.get(key) + "(.*)") && result == false) {
                    result = true;
                    String HeaderGriupName = key.split("_")[0];
                    if(!HeaderGroup.containsKey(HeaderGriupName + "_colStart")){
                        HeaderGroup.put(HeaderGriupName+ "_colStart",String.valueOf(colIndex) );
                    }
                    HeaderGroup.put(HeaderGriupName+ "_colEnd",String.valueOf(colIndex));
                    break;
                }
            }
        }

    }
    public void cellMerge(String cellValue,int colStart, int colEnd,int rowStart,int rowEnd){
        if(firstheaderRow == null){
            firstheaderRow = sheet.createRow(rowStart);
        }
        Cell localCell = firstheaderRow.createCell(colStart);
        localCell.setCellValue(cellValue);
        localCell.setCellStyle(lHeaderStyle);
        CellRangeAddress rangeCell = new CellRangeAddress(rowStart , rowEnd, colStart, colEnd);
        sheet.addMergedRegion(rangeCell);
    }
    public void printDatarowBySchema(final List<String> SchemaList, int atRow) {
        int colIndex = 0;
        if (atRow >= 0) {
            intDataRow = atRow;
        } else {
            intDataRow = intHeaderRow + 1;
        }
        int numLoop = rowCount;
        int colIndex2 = 0;

        for (int index = 0; index < numLoop; index++) {
            rowData data = (rowData) rsRows.get(String.valueOf(index));

            colIndex = 0;
            if (GroupTotalCol.get("GroupCut") != null) {
                if (rowAppend.get(String.valueOf(index)) != null) {

                    /** Print Group **/

                    colIndex2 = 0;
                    row = this.sheet.createRow(intDataRow++);
                    isRequireFormatGroupTotal = true;
                    RowTotalBuffer = 0;
                    for (String key : SchemaList) {
                        whiteDetails(rsTotalRows.get((String) rowAppend.get(String.valueOf(index))).rsColumns.get(key), colIndex2++, key);
                    }
                    whiteDetailsRowTotal(RowTotalBuffer, colIndex2++);
                    isRequireFormatGroupTotal = false;
                    row = this.sheet.createRow(intDataRow++);

                    /** End Print Group **/

                } else {
                    row = this.sheet.createRow(intDataRow++);
                }
            } else {
                row = this.sheet.createRow(intDataRow++);
            }
            RowTotalBuffer = 0;
            for (String key : SchemaList) {

                whiteDetails(data.rsColumns.get(key), colIndex++, key);
            }
            whiteDetailsRowTotal(RowTotalBuffer, colIndex++);
        }
        if (GroupTotalCol.get("GroupCut") != null) {

            colIndex = 0;
            row = this.sheet.createRow(intDataRow++);
            isRequireFormatGroupTotal = true;
            RowTotalBuffer = 0;
            for (String key : localSchemaList) {
                whiteDetails(rsTotalRows.get((String) rowAppend.get(String.valueOf(numLoop))).rsColumns.get(key), colIndex++, key);
            }
            whiteDetailsRowTotal(RowTotalBuffer, colIndex++);
            isRequireFormatGroupTotal = false;
        }

        if (setGrandTotal) {
            /**Grand total All**/
            colIndex = 0;
            row = this.sheet.createRow(intDataRow++);
            isRequireFormatStyle = true;
            RowTotalBuffer = 0;
            for (String key : SchemaList) {

                String value = "";
                if (GroupTotalCol.containsKey("GrantotalCol")) {
                    if (key.matches((String) GroupTotalCol.get("GrantotalCol"))) {
                        value = (String) GroupTotalCol.get("GrantotalTitle");
                        System.err.printf("1-value: %s \n",value);
                        whiteDetails(value, colIndex++, key);
                    } else {
                        value = GrandTotalVal.get(key) == null ? "" : String.valueOf(GrandTotalVal.get(key));
                        System.err.printf("2-value: %s \n",value);
                        whiteDetails(value, colIndex++, key);
                    }
                } else {
                    value = GrandTotalVal.get(key) == null ? "" : String.valueOf(GrandTotalVal.get(key));
                    System.err.printf("3-value: %s \n",value);
                    whiteDetails(value, colIndex++, key);
                }
            }
            whiteDetailsRowTotal(RowTotalBuffer, colIndex++);
            isRequireFormatStyle = false;
            /**End Grand total All**/
        }
    }

    private void whiteDetailsRowTotal(double rowTotalBuffer, int colIndex) {
        if(this.rowSummary == false) return;
        Cell cell = row.createCell(colIndex);
        setBorder(cell);
        if (rowTotalBuffer <= 0) {
            cell.setCellValue("-");
            setTextDelailAlign(cell, HorizontalAlignment.CENTER);
        } else {
            cell.setCellValue(rowTotalBuffer);
            setSingleDigit(cell);
        }
        if (isRequireFormatStyle) {
            setGrandTotalStyle(cell);
        }
        if (isRequireFormatGroupTotal) {
            setGroupTotalStyle(cell);
        }

    }


    public void setLocalSchemaList(List<String> value) {
        localSchemaList = value;
        //setHeaderGroupColumn();

    }
    public void setHeaderGroupColumn() {

        //setHeaderGroupColumn();
        int colIndex =0;
        for(String key:localSchemaList){
            setHeaderGroupColumn(key,colIndex++);
        }
    }
    public void groupTotalLast() {
        System.out.println("groupTotalLast-rowCount :" + rowCount + ", lastCut : " + lastCut);
//        String groupCode = lastGroupCode.substring(0, 3);
        String groupCode = lastCut;
//        String groupCodeDesc = "Total " + lastGroupCode.substring(0, 3);
        String groupCodeDesc = "Total " + lastCut;
        String GroupCutCode = groupCode;
        String groupsuffix = (String) GroupTotalCol.get("GroupSuffix");
        groupCode = groupCode + groupsuffix;
        //rowAppend.put(String.valueOf(rowCount),lastCut);



        /** Add New Row for Group Total **/
        rowAppend.put(String.valueOf(rowCount), groupCode);
        rsTotalRows.put(groupCode, new rowData());
        if (localSchemaList != null) {
            for (String strKeyMark : localSchemaList) {
                if (strKeyMark.equalsIgnoreCase((String) GroupTotalCol.get("GroupCut"))) {
                    rsTotalRows.get(groupCode).addColumn(strKeyMark, "");
                } else if (strKeyMark.equalsIgnoreCase((String) GroupTotalCol.get("GroupCodeTile"))) {
                    rsTotalRows.get(groupCode).addColumn(strKeyMark, groupCode);
                } else if (strKeyMark.equalsIgnoreCase((String) GroupTotalCol.get("GroupCodeDescription"))) {
                    rsTotalRows.get(groupCode).addColumn(strKeyMark, groupCodeDesc);
                } else {
                    double valTemp = GroupTotalVal.get(strKeyMark) != null ? GroupTotalVal.get(strKeyMark) : 0;
                    rsTotalRows.get(groupCode).addColumn(strKeyMark, valTemp);
                }

            }
        }
        /** End Add New Row for Group Total **/
    }

    public void updateGroupTotal() {
        GroupTotalCutVal.put(lastCut, GroupTotalVal);
        GroupTotalClearValue();
    }

    private void workwithData(Schema.Entry entry) {
        String cName = entry.getName();
        String cType = entry.getType().name();
        Object cValue = null;

        switch (cType.toUpperCase()) {
            case "STRING":
                cValue = dataRows.getString(cName);
                break;
            case "BOOLEAN":
                cValue = String.valueOf(dataRows.getBoolean(cName));
                break;
            case "BYTES":
                cValue = dataRows.getBytes(cName).toString();
                break;
            case "DOUBLE":
                cValue = dataRows.getDouble(cName);
                break;
            case "FLOAT":
                cValue = String.valueOf(dataRows.getFloat(cName));
                break;
            case "INT":
                cValue = String.valueOf(dataRows.getInt(cName));
                break;
            case "LONG":
                cValue = String.valueOf(dataRows.getLong(cName));
                break;
            default:
                cValue = String.format("%s", dataRows.get(List.class, cName).toString());
                break;
        }

        ((rowData) rsRows.get(String.valueOf(rowCount))).addColumn(cName, cValue);
        colCount++;

    }

    private void GroupTotalClearValue() {
        GroupTotalVal = new HashMap<String, Double>();
        GroupTotalEmpVal = new HashMap<String, Double>();
        GroupTotalTempVal = new HashMap<String, Double>();
        for (String key : GroupTotalEmpVal.keySet()) {
            GroupTotalEmpVal.put(key, aDoubleTryParse("0"));
        }
        for (String key : GroupTotalTempVal.keySet()) {
            GroupTotalTempVal.put(key, aDoubleTryParse("0"));
        }
        for (String key : GroupTotalVal.keySet()) {
            GroupTotalVal.put(key, aDoubleTryParse("0"));
        }
    }

    private void GroupTotalAddValue(String key, double numValue, int currentColIndex) {
        double tempValue = 0;

        if (formatColumns.containsKey(key)) {
            if (GroupTotalVal != null) {
                if (GroupTotalVal.containsKey(key)) {
                    tempValue = GroupTotalVal.get(key) + numValue;
                    //GroupTotalVal.remove(key);
                    GroupTotalVal.put(key, tempValue);
                } else {
                    GroupTotalVal.put(key, numValue);
                }
            }
        }
    }

    public void appendGrandTotal(String colName, Double value) {

        if (formatColumns.containsKey(colName)) {
            double tempValue = 0.0;
            if (GrandTotalVal.get(colName) != null) {
                tempValue = (double) GrandTotalVal.get(colName) + value;
            } else {
                tempValue = value;
            }
            GrandTotalVal.put(colName, tempValue);
        }
    }

    public void setBorder(Cell cell) {
        if (lDetailsStyle == null) {
            lDetailsStyle = (XSSFCellStyle) wb.createCellStyle();
        }
        lDetailsStyle.cloneStyleFrom(cell.getCellStyle());
        XSSFCellStyle stl = lDetailsStyle;
        stl.setBorderBottom(BorderStyle.THIN);
        stl.setBorderTop(BorderStyle.THIN);
        stl.setBorderRight(BorderStyle.THIN);
        stl.setBorderLeft(BorderStyle.THIN);
        cell.setCellStyle(stl);
    }

    public void setTextAlign(Cell cell, HorizontalAlignment textAlign) {
        if (lHeaderStyle == null) {
            lHeaderStyle = (XSSFCellStyle) wb.createCellStyle();
        }
        lHeaderStyle.cloneStyleFrom(cell.getCellStyle());
        XSSFCellStyle stl = lHeaderStyle;
        stl.setAlignment(textAlign);
        cell.setCellStyle(stl);

    }


    public void setGrandTotalStyle(Cell cell) {
        XSSFCellStyle lDetailsGrandTotalStyle = (XSSFCellStyle) wb.createCellStyle();
        lDetailsGrandTotalStyle.cloneStyleFrom(cell.getCellStyle());
        XSSFCellStyle stl = lDetailsGrandTotalStyle;

        if (colorCodeGrandTotalStyle == null) {
            try {
                colorCodeGrandTotalStyle = new XSSFColor((byte[]) Hex.decodeHex("a8d08d"), null);
            } catch (DecoderException e) {
                e.printStackTrace();
            }
        }
        if (colorCodeGrandTotalStyle != null) {
            stl.setFillForegroundColor(colorCodeGrandTotalStyle);
            stl.setFillBackgroundColor(colorCodeGrandTotalStyle);
            stl.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (fnt2 == null) {
            fnt2 = wb.createFont();
            fnt2.setBold(true);
        }
        stl.setFont(fnt2);
        stl.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0_);[Red](#,##0)"));

        switch (cell.getCellType()) {
            case NUMERIC:
                stl.setAlignment(HorizontalAlignment.RIGHT);
                break;
            default:
                stl.setAlignment(HorizontalAlignment.LEFT);
                break;
        }
        if (cell.getCellType() == CellType.STRING) {
            if (cell.getStringCellValue().matches("-")) {
                stl.setAlignment(HorizontalAlignment.CENTER);
            }
        }
        cell.setCellStyle(stl);

    }

    private void setGroupTotalStyle(Cell cell) {
//        if (lDetailsGroupTotalStyle == null) {
        XSSFCellStyle  lDetailsGroupTotalStyle = (XSSFCellStyle) wb.createCellStyle();

//        }
        lDetailsGroupTotalStyle.cloneStyleFrom(cell.getCellStyle());
        XSSFCellStyle stl = lDetailsGroupTotalStyle;
        //XSSFColor colorCode = null;
        if (colorCodeGroupTotalStype == null) {
            try {
                colorCodeGroupTotalStype = new XSSFColor((byte[]) Hex.decodeHex("8eaadb"), null);
            } catch (DecoderException e) {
                e.printStackTrace();
            }
        }
        if (colorCodeGroupTotalStype != null) {
            //XSSFColor myColor = new XSSFColor(new java.awt.Color(242, 220, 219)); // #f2dcdb
            stl.setFillForegroundColor(colorCodeGroupTotalStype);
            stl.setFillBackgroundColor(colorCodeGroupTotalStype);
            stl.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (fnt2 == null) {
            fnt2 = wb.createFont();
            fnt2.setBold(true);
        }
        stl.setFont(fnt2);
        //stl.setAlignment(HorizontalAlignment.CENTER);
        switch (cell.getCellType()) {
            case NUMERIC:
                stl.setAlignment(HorizontalAlignment.RIGHT);
                break;
            default:
                stl.setAlignment(HorizontalAlignment.LEFT);
                break;
        }
        if (cell.getCellType() == CellType.STRING) {
            //this.setTextAlign(cell,HorizontalAlignment.CENTER);
            if (cell.getStringCellValue().matches("-")) {
                //setTextDelailAlign(cell,HorizontalAlignment.CENTER);
                stl.setAlignment(HorizontalAlignment.CENTER);
            }
        }
        cell.setCellStyle(stl);
    }

    public void setTextDelailAlign(Cell cell, HorizontalAlignment textAlign) {
        if (lDetailsStyleCenter == null) {
            lDetailsStyleCenter = (XSSFCellStyle) wb.createCellStyle();
            lDetailsStyleCenter.cloneStyleFrom(cell.getCellStyle());
        }
        lDetailsStyleCenter.cloneStyleFrom(cell.getCellStyle());

        //XSSFCellStyle stl = (XSSFCellStyle) wb.createCellStyle();
        XSSFCellStyle stl = lDetailsStyleCenter;
        stl.cloneStyleFrom(cell.getCellStyle());
        stl.setAlignment(textAlign);
        cell.setCellStyle(stl);

    }

    public void setCurrencyFormat(Cell cell) {
        if (shortCurrencyFormat == 0) {
            shortCurrencyFormat = wb.getCreationHelper().createDataFormat().getFormat("#,##0.00");
        }
        if (lDetailsStyle == null) {
            lDetailsStyle = (XSSFCellStyle) wb.createCellStyle();
        }
        lDetailsStyle.cloneStyleFrom(cell.getCellStyle());
        XSSFCellStyle stl = lDetailsStyle;
        stl.setDataFormat(shortCurrencyFormat);
        cell.setCellStyle(stl);
    }
    private short shortQtyFormat=0;
    private XSSFCellStyle lDetailsStyleQty = null;
    public void setQtyFormat(Cell cell) {
        if(shortQtyFormat == 0) {
           shortQtyFormat = HSSFDataFormat.getBuiltinFormat("#,##0_);[Red](#,##0)");
        }
        if(lDetailsStyleQty == null){
            lDetailsStyleQty = (XSSFCellStyle) wb.createCellStyle();
        }
        lDetailsStyleQty.cloneStyleFrom(cell.getCellStyle());
        lDetailsStyleQty.setDataFormat(shortQtyFormat);
        lDetailsStyleQty.setAlignment(HorizontalAlignment.RIGHT);
        cell.setCellStyle(lDetailsStyleQty);
    }

    public void setSingleDigit(Cell cell) {

        if (shortSingleDigit == 0) {
            shortSingleDigit = wb.getCreationHelper().createDataFormat().getFormat("#,##0.0");
        }
        if (lDetailsStyle == null) {
            lDetailsStyle = (XSSFCellStyle) wb.createCellStyle();
        }
        lDetailsStyle.cloneStyleFrom(cell.getCellStyle());
        XSSFCellStyle stl = lDetailsStyle;
        stl.setDataFormat(shortSingleDigit);
        cell.setCellStyle(stl);

    }

    public void whiteHeader(String strText, int ColIndex) {
        Cell cell = row.createCell(ColIndex);

        cell.setCellValue(strText);

        if (lHeaderStyle == null) {
            lHeaderStyle = (XSSFCellStyle) wb.createCellStyle();

            if (fnt == null) {
                fnt = wb.createFont();
            }
            if (this.setHeaderBold) fnt.setBold(true);
            if (this.setHeaderBold && this.setHeaderColor) fnt.setColor(IndexedColors.WHITE.getIndex());
            lHeaderStyle.setFont(fnt);
            if (this.setHeaderColor) {
                if (colorCodeHeader == null) {
                    try {
                        colorCodeHeader = new XSSFColor((byte[]) Hex.decodeHex("0070c0"), null);
                    } catch (DecoderException e) {
                        e.printStackTrace();
                    }
                }
                lHeaderStyle.setFillBackgroundColor(colorCodeHeader);
                lHeaderStyle.setFillForegroundColor(colorCodeHeader);
                lHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }

        XSSFCellStyle stl = lHeaderStyle;

        cell.setCellStyle(stl);

        if (this.setHeaderBorder) setBorder(cell);
        if (this.setHeaderRowHeight) row.setHeight((short) -1);
        if (setHeaderTextCenter) setTextAlign(cell, HorizontalAlignment.CENTER);
    }
    public void testCell(){
        Row cRow = sheet.createRow(0);
        Cell cCell = cRow.createCell(0);

        setQtyFormat(cCell);

        cCell.setCellValue(decimalFormat.format(1389.00));

        Cell cCell2 = cRow.createCell(1);

        setCurrencyFormat(cCell2);
        cCell2.setCellValue(123456789.0);

        Cell cCell3 = cRow.createCell(2);

        setQtyFormat(cCell3);
        cCell3.setCellValue(123456789.0);

        Cell cCell4 = cRow.createCell(3);

        setSingleDigit(cCell4);
        cCell4.setCellValue(123456789.0);

    }
    public void whiteDetails( Object objectValue, int colIndex, String colName) {
        Cell cell = row.createCell(colIndex);
        //System.out.println("Print Row: " + String.valueOf(objectValue));
        setBorder(cell);

        if (!formatColumns.isEmpty()) {
            int formatNum = (int) (formatColumns.get(colName) != null ? formatColumns.get(colName) : 0);

            if (formatNum == 1) {
                double tempValue = 0.0;
                try {
                    if(objectValue == null ){
                        tempValue = 0.0;
                    }else {
                        tempValue = new Double(String.valueOf(objectValue));
                    }
                }catch(Exception e){
                    if(objectValue == null ){
                        tempValue = Double.parseDouble("0.0");
                    }else if(((String)objectValue).matches("null")){
                        tempValue = Double.parseDouble("0.0");
                    }
                }

                if(tempValue >0.0) {
                    this.setQtyFormat(cell);
                    cell.setCellValue(tempValue);

                }else {
                    cell.setCellValue("-");
                    setTextDelailAlign(cell, HorizontalAlignment.CENTER);
                }
            } else if (formatNum == 2) {
                this.setCurrencyFormat(cell);
                cell.setCellValue((double) objectValue);
            } else if (formatNum == 3) {
                SimpleDateFormat myFormat = new SimpleDateFormat("dd-MMM-yyyy");
                try {
                    String strValue = myFormat.format(new SimpleDateFormat("EEE MMM dd hh:mm:ss zz yyyy").parse((String) objectValue));
                    cell.setCellValue(strValue);
                } catch (ParseException e) {

                    System.out.println("Can't convert date format : " + (String) objectValue);
                    System.err.println("Error: " + (String) e.getMessage());
                    cell.setCellValue((String) objectValue);
                    //e.printStackTrace();
                }
            } else if (formatNum == 4) {
                this.setSingleDigit(cell);
                double tempValue = 0.0;
                try {
                    tempValue = Double.parseDouble(String.valueOf(objectValue));
                } catch (Exception e) {
                    tempValue = 0.0;
                }

                cell.setCellValue(tempValue);
            } else if (formatNum == 5) {
                this.setSingleDigit(cell);
                double tempValue = 0.0;
                try {
                    tempValue = Double.parseDouble(String.valueOf(objectValue));
                } catch (Exception e) {
                    tempValue = 0.0;
                }
                if (tempValue > 0) {
                    cell.setCellValue(tempValue);
                } else {
                    cell.setCellValue("-");
                    setTextDelailAlign(cell, HorizontalAlignment.CENTER);
                }
            } else {
                cell.setCellValue(String.valueOf(objectValue));
            }
        } else {
            cell.setCellValue(String.valueOf(objectValue));
        }
        if (isRequireFormatStyle) {
            setGrandTotalStyle(cell);
        }

        if (isRequireFormatGroupTotal) {
            setGroupTotalStyle(cell);
        }
        if (RowTotalColStart <= colIndex) {
            RowTotalBuffer = RowTotalBuffer + aDoubleTryParse(objectValue != null ? objectValue.toString() : "0");
            ;
        }
    }

    private static double aDoubleTryParse(String DigitValue) {
        double tempValue = 0;
        try {
            tempValue = Double.parseDouble(DigitValue);
        } catch (Exception e) {
            tempValue = 0;
        }
        return tempValue;
    }

    public static class cTotal {
        public static Map<String, Double> dValue = new HashMap<String, Double>();
        public static Map<String, Integer> iValue = new HashMap<String, Integer>();
        public static Map<String, String> sValue = new HashMap<String, String>();

        public void clearValue() {
            dValue.clear();
            iValue.clear();
            sValue.clear();
        }

        public void addValue(String name, double value) {
            if (dValue.containsKey(name)) {
                double ldValue = dValue.get(name);
                dValue.put(name, ldValue + value);
            } else {
                dValue.put(name, value);
            }
        }

        public void addValue(String name, int value) {
            if (iValue.containsKey(name)) {
                int liValue = iValue.get(name);
                iValue.put(name, liValue + value);
            } else {
                iValue.put(name, value);
            }

        }

        public void addValue(String name, String value) {
            if (sValue.containsKey(name)) {
                String lsValue = sValue.get(name);
                sValue.put(name, lsValue + "," + value);
            } else {
                sValue.put(name, value);
            }

        }
    }
}
