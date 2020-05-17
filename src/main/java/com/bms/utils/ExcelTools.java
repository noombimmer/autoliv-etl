package com.bms.utils;

import com.sun.java.swing.plaf.windows.WindowsButtonListener;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.talend.sdk.component.api.record.Record;
import org.talend.sdk.component.api.record.Schema;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;

public class ExcelTools {
    private static File fileInstanced = null;
    private static FileInputStream fIP = null;
    private static String lFileName;
    private static String lSheetName="Sheet1";
    private static String prevSheetName="";
    private static Workbook wb = null;
    private static Sheet sheet = null;
    private int rowAccessWindowSize = SXSSFWorkbook.DEFAULT_WINDOW_SIZE;// used in auto flush
    private boolean appendWorkbook = false;
    private boolean appendSheet = false;
    private boolean recalculateFormula = false;

    public boolean setHeaderColor = true;
    public boolean setHeaderBorder = true;
    public boolean setHeaderBold = true;
    public boolean setHeaderRowHeight = true;
    public boolean setHeaderTextCenter = true;
    private static int intHeaderRow = 0;
    private static int intDataRow = 0;
    private static int colCount = 0;
    private static int rowCount = 0;
    private static Record dataRows;
    private static Row row;
    private static java.util.Map<String,Object> globalColumns = new java.util.HashMap<String,Object>();
    private static java.util.Map<String,Object> formatColumns = new java.util.HashMap<String,Object>();
    //sheetObject
    public java.util.Map<String,sheetObject> dataContents = new java.util.HashMap<String,sheetObject>();
    private static int globalColIndex = 0;

    public void addContent(sheetObject data,String name){
        dataContents.put(name,data);
    }

    public ExcelTools(String fileName){
        this.lFileName = fileName;
        fileInstanced = new File(this.lFileName);
        if(fileInstanced.exists()){
            try {
                this.fIP = new FileInputStream(this.fileInstanced);
                wb = new XSSFWorkbook(fIP);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }else{
            fileInstanced.delete();
            wb = new SXSSFWorkbook(rowAccessWindowSize);
            sheet = wb.createSheet(lSheetName);
        }
    }
    public void reloadFile(){
        if(this.lFileName.isEmpty()) return;
        if(wb == null) return;
        if(sheet == null) return;
        if(this.fIP != null){
            try {
                this.fIP.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        if(fileInstanced.exists()){
            try {
                this.fIP = new FileInputStream(this.fileInstanced);
                wb = new XSSFWorkbook(fIP);
                sheet = wb.getSheet(lSheetName);
                lDetailsStyle = null;
                lHeaderStyle = null;
                shortCurrencyFormat = 0;
                shortQtyFormat = 0;
            } catch (IOException e) {
                e.printStackTrace();
            }
        }else{
            return;
        }
    }
    public ExcelTools(String fileName,String sheetName){
        this.lFileName = fileName;
        if(!sheetName.isEmpty() && sheetName.length() > 0 && sheetName != null){
            this.lSheetName = sheetName;
        }
        fileInstanced = new File(this.lFileName);
        if(fileInstanced.exists()){
            try {
                this.fIP = new FileInputStream(this.fileInstanced);
                wb = new XSSFWorkbook(fIP);
                sheet = wb.getSheet(lSheetName);
                if(sheet != null){
                    wb.removeSheetAt(wb.getSheetIndex(lSheetName));
                    sheet = wb.createSheet(lSheetName);
                }else{
                    sheet = wb.createSheet(lSheetName);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }else{
            fileInstanced.delete();
            wb = new SXSSFWorkbook(rowAccessWindowSize);
            sheet = wb.createSheet(lSheetName);
        }
    }
    public ExcelTools(){

    }
    public void writeExcel(OutputStream outputStream) throws Exception {
        wb.write(outputStream);

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
        shortQtyFormat = 0;
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
    public void closedFile(String xlsFilename,boolean createDir){
        System.out.print("\b\b\b\b Save File : " + xlsFilename + "\r\n");
        if (createDir) {
            File file = new File(xlsFilename);
            File pFile = file.getParentFile();
            if (pFile != null && !pFile.exists()) {
                pFile.mkdirs();
            }
        }
        FileOutputStream fileOutput = null;
        try {
            fileOutput = new FileOutputStream(xlsFilename);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        if (appendWorkbook && appendSheet && recalculateFormula) {
            evaluateFormulaCell();
        }
        try {
            wb.write(fileOutput);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            fileOutput.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public void createSheet(){

        this.sheet = wb.getSheet(this.lSheetName);
        if(this.sheet == null){
            wb.createSheet(this.lSheetName);
            System.out.printf("Create Sheet ...%s\n",lSheetName);
        }else{
            System.out.printf("Open Sheet ...%s\n",lSheetName);
        }

    }
    public void setAutoSizeCol(){
        if(this.sheet != null){
            System.out.print("\b\b\b\b Set AutoSize Columns\r\n");
            int colIndex = 0;
            for(String key:localSchemaList){
                //System.out.println("Set column :" + key + " Autosize " + String.valueOf(colIndex));
                this.sheet.autoSizeColumn(colIndex++);

            }
        }
    }

    public void clearData(){
        colCount = 0;
        rowCount = 0;
        dataRows = null;
        intHeaderRow = 0;
        intDataRow = 0;
        row = null;
        if(globalColumns !=null) globalColumns.clear();
        if(formatColumns !=null) formatColumns.clear();
        if(rsRows !=null) rsRows.clear();
        if(localSchemaList !=null) localSchemaList.clear();
        globalColIndex=0;
        lHeaderStyle =null;
        lDetailsStyle = null;

        File fileInstanced = null;
        FileInputStream fIP = null;
        lFileName = "";

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
        System.out.printf("Clear Data Class ...%s\n",lSheetName);
        lSheetName="Sheet1";
    }
    public static void setColumnFormat(String colName,int formatNum){
        formatColumns.put(colName,formatNum);
    }

    public void setLocalSchemaList(List<String> config) {
        localSchemaList = config;
    }
    public void readLastObject(){
        if(rowCount > 0 ){
            sheetObject data = new sheetObject();
            data.rowCount = rowCount;
            data.ColList = localSchemaList;
            data.colCount = colCount;
            data.sheet = sheet;
            data.sheetName = lSheetName;
            data.wb = wb;
            data.fileName = lFileName;
            data.formatColumns = formatColumns;
            addContent(data,lSheetName);
        }
    }
    public void resetRow() {
        System.out.printf("WB Class ...%s\n",wb == null? "null": wb.toString());
        System.out.printf("Sheet Class ...%s\n",sheet == null? "null": sheet.toString());
        System.out.printf("Row Class ...%s\n",row == null? "null": row.toString());
        System.out.printf("rowCount  ...%d\n",rowCount);
        System.out.printf("intHeaderRow  ...%d\n",intHeaderRow);
        System.out.printf("intDataRow  ...%d\n",intDataRow);
        System.out.printf("colCount  ...%d\n",colCount);
        System.out.printf("prevSheetName  ...%s\n",prevSheetName);

        if(rowCount > 0 && !prevSheetName.equalsIgnoreCase(lSheetName) ){
            sheetObject data = new sheetObject();
            data.rowCount = rowCount;
            data.ColList = localSchemaList;
            data.colCount = colCount;
            data.sheet = sheet;
            data.sheetName = prevSheetName;
            data.wb = wb;
            data.fileName = lFileName;
            addContent(data,prevSheetName);
        }else{
            prevSheetName = lSheetName;
            System.out.printf("Deposite...%s\n",prevSheetName);
        }
        this.rowCount = 0;
        this.intHeaderRow =0;
        this.intDataRow = 0;
        this.colCount = 0;
        sheet = null;

        System.out.printf("Reset Data Class ...%s\n",lSheetName);
    }

    public String getSheetName() {
        return lSheetName;
    }

    public void resetClass() {
        dataContents.clear();
        dataContents = new java.util.HashMap<String,sheetObject>();
        rowCount =0;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public void setwb(Workbook wb) {
        this.wb = wb;
    }

    public void setDataContents(sheetObject data) {
        setwb(data.wb);
        setSheet(data.sheet);
        rowCount = data.rowCount;
        lSheetName = data.sheetName;
        lFileName = data.fileName;

    }

    public static class sheetObject{
        public String sheetName;
        public int rowCount;
        public int colCount;
        public Workbook wb =null;
        public Sheet sheet =null;
        public List<String> ColList;
        public String fileName;
        public Map<String, Object> formatColumns;

        public void setSheetName(String name){
            sheetName = name;
        }

        public void setRowCout(int count){
            rowCount = count;
        }
        public void setColCout(int count){
            colCount = count;
        }

    }

    public static class rowData {
        private java.util.Map<String,Object> rsColumns = new java.util.HashMap<String,Object>();
        //private int colIndex = 0;

        public rowData(){
        }
        public void addColumn(String colName,Object colValue){
            if(rsColumns.get(colName) == null){
                rsColumns.put(colName,colValue);
            }
            addGlobalColumn(colName);
        }
    }
    private java.util.Map<String,rowData> rsRows = new java.util.HashMap<String,rowData>();
    public static void addGlobalColumn(String colName){
        if(globalColumns.get(colName) == null){
            globalColumns.put(colName,globalColIndex++);
        }
    }
    public void getDataFromRecord(Record datainput){
        colCount = 0;
        dataRows = datainput;
        List<Schema.Entry> arrayRec = datainput.getSchema().getEntries();

        //row = this.sheet.createRow(rowCount);
        rsRows.put(String.valueOf(rowCount),new rowData());
        //System.out.println("Row size: " + arrayRec.size());

        arrayRec.forEach(this::workwithData);
        //arrayRec.get(0).getName();
        rowCount++;
        //System.out.print("\b\b\b\b Count Rows: "+rowCount+"\r");
    }
    public void printHeader(int atRow){
        int colIndex = 0;
        if(atRow >= 0 ){
            intHeaderRow = atRow;
        }
        for (String key: globalColumns.keySet()){
            if(row == null){
                if(sheet ==null ){
                    if(wb == null){
                        this.reloadFile();
                    }
                    this.createSheet();
                    row = this.sheet.createRow(intHeaderRow);
                }else{
                    row = this.sheet.createRow(intHeaderRow);
                }

            }
            whiteHeader(key,colIndex++);
        }
    }
    public static List<String> localSchemaList;
    public void printHeaderBySchema(final List<String> SchemaList,int atRow){
        localSchemaList = SchemaList;
        int colIndex = 0;
        if(atRow >= 0 ){
            intHeaderRow = atRow;
        }
        for (String key: SchemaList){
            if(row == null){
                row = this.sheet.createRow(intHeaderRow);
            }
            whiteHeader(key,colIndex++);
        }
    }

    public void printDatarow(int atRow){
        int colIndex = 0;
        if(atRow >= 0 ){
            intDataRow = atRow;
        }else{
            intDataRow = intHeaderRow + 1;
        }
        for(String intRowNum:rsRows.keySet()){
            rowData data =(rowData) rsRows.get(intRowNum);
            colIndex=0;
            row = this.sheet.createRow(intDataRow++);
            for (String key: globalColumns.keySet()){
                whiteDetails((String) data.rsColumns.get(key),colIndex++);
            }
        }
    }
    public void printDatarowBySchema(final List<String> SchemaList,int atRow){
        int colIndex = 0;
        if(atRow >= 0 ){
            intDataRow = atRow;
        }else{
            intDataRow = intHeaderRow + 1;
        }
        int numLoop = rowCount;
        for(int index = 0;index < numLoop;index++){
            rowData data =(rowData) rsRows.get(String.valueOf(index));
            colIndex=0;
            row = this.sheet.createRow(intDataRow++);
            for (String key: SchemaList){
                whiteDetails( data.rsColumns.get(key),colIndex++,key);
            }
        }
        /*
        for(String intRowNum:rsRows.keySet()){
            rowData data =(rowData) rsRows.get(intRowNum);
            colIndex=0;
            row = this.sheet.createRow(intDataRow++);
            for (String key: SchemaList){
                whiteDetails( data.rsColumns.get(key),colIndex++,key);
            }
        }
         */
    }
    CellStyle lHeaderStyle = null;
    CellStyle lDetailsStyle = null;
    private void workwithData(Schema.Entry entry) {
        String cName = entry.getName();
        String cType = entry.getType().name();
        Object cValue =null;
        if(cName.contentEquals("Row")){
            //System.out.println("Data Type: " + cType.toUpperCase());
        }
        switch (cType.toUpperCase()){
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
                cValue = String.format("%s",dataRows.get(List.class,cName).toString());
                break;
        }

        ((rowData)rsRows.get(String.valueOf(rowCount))).addColumn(cName,cValue);
        /*
        if(rowCount == 0){
            whiteHeader(cName);
        }else{
            whiteDetails("[" + colCount + "]::" +cName + ":::"+ cValue);
        }
        */

        colCount++;

    }
    public void setBorder(Cell cell){
        if(lDetailsStyle == null){
            lDetailsStyle = wb.createCellStyle();
        }
        CellStyle stl=lDetailsStyle;
        stl.setBorderBottom(BorderStyle.THIN);
        stl.setBorderTop(BorderStyle.THIN);
        stl.setBorderRight(BorderStyle.THIN);
        stl.setBorderLeft(BorderStyle.THIN);
        cell.setCellStyle(stl);
    }
    public void setTextAlign(Cell cell,HorizontalAlignment textAlign){
        if(lHeaderStyle == null){
            lHeaderStyle = wb.createCellStyle();
        }
        CellStyle stl=lHeaderStyle;
        stl.setAlignment(textAlign);
        cell.setCellStyle(stl);

    }

    public short shortCurrencyFormat = 0;
    public short shortQtyFormat = 0;
    public void setCurrencyFormat(Cell cell){
        if(shortCurrencyFormat == 0){
            shortCurrencyFormat = wb.getCreationHelper().createDataFormat().getFormat("#,##0.00");
        }
        if(lDetailsStyle == null){
            lDetailsStyle = wb.createCellStyle();
        }
        CellStyle stl=lDetailsStyle;
        stl.setDataFormat(shortCurrencyFormat);
        cell.setCellStyle(stl);
    }
    public void setQtyFormat(Cell cell){
        if(shortQtyFormat == 0){
            shortQtyFormat = wb.getCreationHelper().createDataFormat().getFormat("#,##0");
        }
        if(lDetailsStyle == null){
            lDetailsStyle = wb.createCellStyle();
        }
        CellStyle stl=lDetailsStyle;
        stl.setDataFormat(shortQtyFormat);
        cell.setCellStyle(stl);
    }

    public void whiteHeader(String strText){
        Cell cell = row.createCell(colCount);
        cell.setCellValue(strText);

        if(lHeaderStyle == null) {
            lHeaderStyle = wb.createCellStyle();

            Font fnt = wb.createFont();
            if(this.setHeaderBold) fnt.setBold(true);
            if(this.setHeaderBold && this.setHeaderColor)fnt.setColor(IndexedColors.WHITE.getIndex());
            lHeaderStyle.setFont(fnt);
            if(this.setHeaderColor) {
                lHeaderStyle.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                lHeaderStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                lHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }
        CellStyle stl=lHeaderStyle;
        cell.setCellStyle(stl);
        if(this.setHeaderBorder) setBorder(cell);
        if(this.setHeaderRowHeight) row.setHeight((short)-1);
    }
    public void whiteHeader(String strText,int ColIndex){
        Cell cell = row.createCell(ColIndex);

        cell.setCellValue(strText);

        if(lHeaderStyle == null) {
            lHeaderStyle = wb.createCellStyle();

            Font fnt = wb.createFont();
            if(this.setHeaderBold) fnt.setBold(true);
            if(this.setHeaderBold && this.setHeaderColor)fnt.setColor(IndexedColors.WHITE.getIndex());
            lHeaderStyle.setFont(fnt);
            if(this.setHeaderColor) {
                lHeaderStyle.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                lHeaderStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
                lHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
        }
        CellStyle stl=lHeaderStyle;

        cell.setCellStyle(stl);

        if(this.setHeaderBorder) setBorder(cell);
        if(this.setHeaderRowHeight) row.setHeight((short)-1);
        if(setHeaderTextCenter) setTextAlign(cell,HorizontalAlignment.CENTER );
    }

    public void whiteDetails(String strText){
        Cell cell = row.createCell(colCount);
        cell.setCellValue(strText);
    }
    public void whiteDetails(String strText,int colIndex){
        Cell cell = row.createCell(colIndex);
        cell.setCellValue(strText);
        setBorder(cell);

    }
    public void whiteDetails(Object objectValue,int colIndex,String colName){
        Cell cell = row.createCell(colIndex);

        setBorder(cell);
        if(!formatColumns.isEmpty()) {
            int formatNum = (int) (formatColumns.get(colName) != null ?formatColumns.get(colName) :0 ) ;
            if (formatNum == 1) {
                this.setQtyFormat(cell);
                cell.setCellValue((double)objectValue);
            }else if (formatNum == 2) {
                this.setCurrencyFormat(cell);
                cell.setCellValue((double)objectValue);
            }else if (formatNum == 3) {
                SimpleDateFormat myFormat = new SimpleDateFormat("dd-MMM-yyyy");
                try {
                    String strValue = myFormat.format(new SimpleDateFormat("EEE MMM dd hh:mm:ss zz yyyy").parse((String) objectValue));
                    cell.setCellValue(strValue);
                } catch (ParseException e) {
                    String strValue = null;
                    try {
                        strValue = myFormat.format(new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").parse((String) objectValue));
                    } catch (ParseException parseException) {
                        //parseException.printStackTrace();
                        strValue="N/A";
                        System.err.println(e.getMessage());
                    }
                    cell.setCellValue(strValue);

                }
            }else if(formatNum == 4){
                if( objectValue  instanceof Double ){
                    setSingleDigit(cell);
                    cell.setCellValue((double)objectValue);
                }else{
                    try{
                        Double valueTemp = Double.parseDouble(objectValue.toString());
                        setSingleDigit(cell);
                        cell.setCellValue(valueTemp);
                    }catch(Exception e){
                        setSingleDigit(cell);
                        cell.setCellValue(0.00);
                    }
                }
            }
            else{
                cell.setCellValue(String.valueOf(objectValue));
            }
        }else{
            cell.setCellValue(String.valueOf(objectValue));
        }

    }
    private XSSFCellStyle lDetailsStyleSingleDigit =null;
    private void setSingleDigit(Cell cell) {

        if(lDetailsStyleSingleDigit == null){
            lDetailsStyleSingleDigit = (XSSFCellStyle) wb.createCellStyle();
        }
        lDetailsStyleSingleDigit.cloneStyleFrom(cell.getCellStyle());
        lDetailsStyleSingleDigit.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("#,##0.0"));
        lDetailsStyleSingleDigit.setAlignment(HorizontalAlignment.RIGHT);
        cell.setCellStyle(lDetailsStyleSingleDigit);
    }

}
