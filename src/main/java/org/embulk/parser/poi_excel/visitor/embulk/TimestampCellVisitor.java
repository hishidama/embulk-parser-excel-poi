package org.embulk.parser.poi_excel.visitor.embulk;

import java.time.Instant;
import java.time.format.DateTimeParseException;
import java.util.Date;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.PluginTask;
import org.embulk.parser.poi_excel.PoiExcelParserPlugin.TimestampColumnOption;
import org.embulk.parser.poi_excel.visitor.PoiExcelVisitorValue;
import org.embulk.spi.Column;
import org.embulk.spi.type.TimestampType;
import org.embulk.util.config.ConfigMapper;
import org.embulk.util.config.ConfigMapperFactory;
import org.embulk.util.config.units.ColumnConfig;
import org.embulk.util.config.units.SchemaConfig;
import org.embulk.util.timestamp.TimestampFormatter;

public class TimestampCellVisitor extends CellVisitor {

    protected static final ConfigMapper CONFIG_MAPPER;
    static {
        ConfigMapperFactory factory = ConfigMapperFactory.builder().addDefaultModules().build();
        CONFIG_MAPPER = factory.createConfigMapper();
    }

    public TimestampCellVisitor(PoiExcelVisitorValue visitorValue) {
        super(visitorValue);
    }

    @Override
    public void visitCellValueNumeric(Column column, Object source, double value) {
        TimeZone timeZone = getTimestampParser(column).zone();
        Date date = DateUtil.getJavaDate(value, timeZone);
        Instant timestamp = Instant.ofEpochMilli(date.getTime());
        pageBuilder.setTimestamp(column, timestamp);
    }

    @Override
    public void visitCellValueString(Column column, Object source, String value) {
        TimestampFormatter formatter = getTimestampParser(column).formatter();
        Instant timestamp;
        try {
            timestamp = formatter.parse(value);
        } catch (DateTimeParseException e) {
            doConvertError(column, value, e);
            return;
        }
        pageBuilder.setTimestamp(column, timestamp);
    }

    @Override
    public void visitCellValueBoolean(Column column, Object source, boolean value) {
        doConvertError(column, value, new UnsupportedOperationException("unsupported conversion Excel boolean to Embulk timestamp"));
    }

    @Override
    public void visitCellValueError(Column column, Object source, int code) {
        doConvertError(column, code, new UnsupportedOperationException("unsupported conversion Excel Cell error code to Embulk timestamp"));
    }

    @Override
    public void visitValueLong(Column column, Object source, long value) {
        pageBuilder.setTimestamp(column, Instant.ofEpochMilli(value));
    }

    @Override
    public void visitSheetName(Column column) {
        Sheet sheet = visitorValue.getSheet();
        visitSheetName(column, sheet);
    }

    @Override
    public void visitSheetName(Column column, Sheet sheet) {
        doConvertError(column, sheet.getSheetName(), new UnsupportedOperationException("unsupported conversion sheet_name to Embulk timestamp"));
    }

    @Override
    public void visitRowNumber(Column column, int index1) {
        doConvertError(column, index1, new UnsupportedOperationException("unsupported conversion row_number to Embulk timestamp"));
    }

    @Override
    public void visitColumnNumber(Column column, int index1) {
        doConvertError(column, index1, new UnsupportedOperationException("unsupported conversion column_number to Embulk timestamp"));
    }

    @Override
    protected void doConvertErrorConstant(Column column, String value) throws Exception {
        TimestampFormatter formatter = getTimestampParser(column).formatter();
        Instant timestamp = formatter.parse(value);
        pageBuilder.setTimestamp(column, timestamp);
    }

    private static class TimestampInfo {
        private final TimeZone zone;
        private final TimestampFormatter formatter;

        public TimestampInfo(TimeZone zone, TimestampFormatter formatter) {
            this.zone = zone;
            this.formatter = formatter;
        }

        public TimeZone zone() {
            return zone;
        }

        public TimestampFormatter formatter() {
            return formatter;
        }
    }

    private TimestampInfo[] timestampParsers;

    // https://zenn.dev/dmikurube/articles/get-ready-for-embulk-v0-11-and-v1-0
    protected final TimestampInfo getTimestampParser(Column column) {
        if (timestampParsers == null) {
            PluginTask task = visitorValue.getPluginTask();
            SchemaConfig schema = task.getColumns();
            TimestampInfo[] parsers = new TimestampInfo[schema.getColumnCount()];
            int i = 0;
            for (ColumnConfig c : schema.getColumns()) {
                if (c.getType() instanceof TimestampType) {
                    TimestampColumnOption columnOption = CONFIG_MAPPER.map(c.getOption(), TimestampColumnOption.class);
                    String zoneString = columnOption.getTimeZoneId().orElse(task.getDefaultTimeZoneId());
                    TimeZone zone = TimeZone.getTimeZone(zoneString);
                    TimestampFormatter formatter = TimestampFormatter.builder(columnOption.getFormat().orElse(task.getDefaultTimestampFormat()), true) //
                            .setDefaultZoneFromString(zoneString) //
                            .setDefaultDateFromString(columnOption.getDate().orElse(task.getDefaultDate())) //
                            .build();
                    parsers[i] = new TimestampInfo(zone, formatter);
                }
                i++;
            }
            timestampParsers = parsers;
        }
        return timestampParsers[column.getIndex()];
    }
}
