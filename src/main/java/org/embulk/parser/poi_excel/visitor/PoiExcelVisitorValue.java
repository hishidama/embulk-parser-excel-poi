package org.embulk.parser.poi_excel.visitor;

import java.lang.reflect.Method;

import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.bean.PoiExcelSheetBean;
import org.embulk.spi.Column;
import org.embulk.spi.FileInput;
import org.embulk.spi.PageBuilder;
import org.embulk.spi.Schema;

public class PoiExcelVisitorValue {
    private final PluginTask task;
    private final FileInput input;
    private final Sheet sheet;
    private final PageBuilder pageBuilder;
    private final PoiExcelSheetBean sheetBean;
    private PoiExcelVisitorFactory factory;

    public PoiExcelVisitorValue(PluginTask task, Schema schema, FileInput input, Sheet sheet, PageBuilder pageBuilder) {
        this.task = task;
        this.input = input;
        this.sheet = sheet;
        this.pageBuilder = pageBuilder;
        this.sheetBean = new PoiExcelSheetBean(task, schema, sheet);
    }

    public PluginTask getPluginTask() {
        return task;
    }

    private Method hintOfCurrentInputFileNameForLogging;
    private Method optional_orElse;

    public String getFileName() {
        if (hintOfCurrentInputFileNameForLogging == null) {
            try {
                this.hintOfCurrentInputFileNameForLogging = FileInput.class.getMethod("hintOfCurrentInputFileNameForLogging");
                this.optional_orElse = Class.forName("java.util.Optional").getMethod("orElse", Object.class);
            } catch (Exception e) {
                throw new RuntimeException("use Embulk 0.9.12 or later for value=file_name", e);
            }
        }

        try {
            Object fileNameOption = hintOfCurrentInputFileNameForLogging.invoke(input);
            return (String) optional_orElse.invoke(fileNameOption, (String) null);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public PageBuilder getPageBuilder() {
        return pageBuilder;
    }

    public void setVisitorFactory(PoiExcelVisitorFactory factory) {
        this.factory = factory;
    }

    public PoiExcelVisitorFactory getVisitorFactory() {
        return factory;
    }

    public PoiExcelSheetBean getSheetBean() {
        return sheetBean;
    }

    public PoiExcelColumnBean getColumnBean(Column column) {
        return sheetBean.getColumnBean(column);
    }
}
