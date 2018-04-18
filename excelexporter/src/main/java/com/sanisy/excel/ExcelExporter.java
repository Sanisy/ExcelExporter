package com.sanisy.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * exporter the data to the excel
 * Created by Sanisy on 2018/4/18.
 */
@Slf4j
public class ExcelExporter {

    private Class<?> clazz;

    /**
     * the excel file
     */
    private File file;

    private String sheetName;

    //save the title of the target column index
    private TreeMap<String, String> titleMap = new TreeMap<String, String>();

    //save the fieldName of the target column index
    private Map<String, String> fieldNameMap = new HashMap<String, String>();

    //save the field's data's converter of the target column index
    private Map<String, Class<?>> fieldConverterMap = new HashMap<String, Class<?>>();

    //save the field's data's creator of the target column index, this class will offer the data of the target cell
    private Map<String, Class<?>> cellPipelineMap = new HashMap<String, Class<?>>();

    private DefaultColumnDataConverter converter = new DefaultColumnDataConverter();

    public ExcelExporter() {

    }

    public static class Builder{

        /**
         * default excel file name
         */
        private static final String DEFAULT_FILE_NAME = "excel0.xlsx";

        private File file;

        private String fileName = DEFAULT_FILE_NAME;

        /**
         * default excel's sheet name
         */
        private static final String DEFAULT_SHEET_NAME = "sheet0";

        private String sheetName = DEFAULT_SHEET_NAME;

        public Builder() {

        }

        public ExcelExporter build() {
            ExcelExporter excelExporter = new ExcelExporter();
            excelExporter.file = this.file;
            excelExporter.sheetName = sheetName;

            return excelExporter;
        }


        public Builder exportTo(String fileName, String sheetName) {
            int last = fileName.lastIndexOf(".");
            if (last > -1) {
                this.fileName = fileName.substring(0, last) + ".xlsx";
            }

            if (file == null) {
                file = new File(this.fileName);
                try {
                    file.createNewFile();
                } catch (IOException e) {
                    log.error("create file error,fileName={}", fileName, e);
                }
            }

            if (sheetName != null && !"".equals(sheetName.trim())) {
                this.sheetName = sheetName;
            }

            return this;
        }

        public Builder exportTo(File file, String sheetName) {
            if (file != null) {
              this.file = file;
            }

            if (sheetName != null && !"".equals(sheetName.trim())) {
                this.sheetName = sheetName;
            }

            return this;
        }
    }


    /**
     * export data to target file
     *
     * @param dataList
     */
    public void doExport(List<?> dataList) {
        if (dataList == null || dataList.size() == 0) {
            log.info("data source is null or empty, dataList={}", dataList);
            return;
        }

        if (file == null) {
            throw new NullPointerException("the target excel file can not be null, check your exportTo method's parameter");
        }

        getAnnotationInfo(dataList.get(0));
        createExcel(dataList);
    }


    /**
     * load the annotations info of ExcelColumn
     * @param t
     */
    private void getAnnotationInfo(Object t) {
        clazz = t.getClass();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {

            //获取注解
            if (field.isAnnotationPresent(ExcelIgnore.class)) {
                continue;
            }

            if (field.isAnnotationPresent(ExcelColumn.class)) {
                try {
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    //index
                    String index = excelColumn.index();
                    fieldNameMap.put(index, field.getName());
                    //title
                    String title = excelColumn.title();
                    titleMap.put(index, title);

                    //处理handler
                    Class<?> converterClazz = excelColumn.converter();
                    if (ColumnDataConverter.class.isAssignableFrom(converterClazz)) {
                        fieldConverterMap.put(index, converterClazz);
                    }

                    Class<?> dataPipeline = excelColumn.dataPipeline();
                    if (CellDataPipeline.class.isAssignableFrom(dataPipeline)) {
                        cellPipelineMap.put(index, dataPipeline);
                    }

                } catch (Exception e) {
                    log.error("occurs exception when get the annotation field's data", e);
                }
            }
        }
    }

    private void createExcel(List<?> dataList){
        OutputStream outputStream = null;
        try {
            outputStream = new BufferedOutputStream(new FileOutputStream(file));
        } catch (Exception e) {
            log.error("error occurs when open outputStream of the excel file, fileName={}", file.getName(), e);
        }

        int poolSize = 10000;
        SXSSFWorkbook wb = new SXSSFWorkbook(poolSize);
        Sheet sheet = wb.createSheet(sheetName);
        //写title
        Row titleRow = sheet.createRow(0);
        List<Integer> columnIndexList = new ArrayList<Integer>();
        for (String indexObj : fieldNameMap.keySet()) {
            Integer index = Integer.parseInt(indexObj);
            columnIndexList.add(index);
            Cell cell = titleRow.createCell(index);
            cell.setCellValue(titleMap.get(index + ""));
        }

        //
        int total = dataList.size();
        for (int i = 0; i < total; i++) {
            Row row = sheet.createRow(i + 1);
            for (Integer index : columnIndexList) {
                Cell cell = row.createCell(index);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                //the object of the current row
                Object t = dataList.get(i);
                String fieldStr = fieldNameMap.get(index + "");
                try {
                    Class<?> dataPipelineClazz = cellPipelineMap.get(index + "");
                    if (dataPipelineClazz != null && CellDataPipeline.class.isAssignableFrom(dataPipelineClazz)) {//this column's data is customized by user
                        CellDataPipeline instance = (CellDataPipeline)dataPipelineClazz.newInstance();
                        String data = instance.dataOfCell(i + 1);
                        cell.setCellValue(data);
                    }else {
                        Field field = clazz.getDeclaredField(fieldStr);
                        field.setAccessible(true);
                        String cellValue = getCellValue(index, field.get(t));
                        cell.setCellValue(cellValue);
                    }
                } catch (NoSuchFieldException e) {
                    log.error("can not found the field,field={}", fieldStr, e);
                } catch (IllegalAccessException e) {
                    log.error("access permission limited,the field's modifier is private,permission limit,field={}", fieldStr, e);
                } catch (InstantiationException e) {
                    log.error("reflection instance error", e);
                }
            }
        }

        try {
            wb.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.flush();
                    outputStream.close();
                } catch (IOException e) {
                    log.error("excel file output stream close error,fileName={}", file.getName(), e);
                }
            }

            wb.dispose();
        }
    }

    private String getCellValue(Integer index, Object data) {
        Class<?> converterClass = fieldConverterMap.get(index);
        try {
            if (converterClass == null) {
                return converter.convert(data);
            }

            ColumnDataConverter instance = (ColumnDataConverter)converterClass.newInstance();
            return instance.convert(data);
        } catch (Exception e) {
            log.error("data convert to String error, data={}", data, e);
        }

        return "";
    }
}
