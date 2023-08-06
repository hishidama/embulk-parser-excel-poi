package org.embulk.parser.poi_excel.visitor;

import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.embulk.parser.poi_excel.PoiExcelColumnValueType;
import org.embulk.parser.poi_excel.bean.PoiExcelColumnBean;
import org.embulk.parser.poi_excel.bean.record.PoiExcelRecord;
import org.embulk.parser.poi_excel.util.PoiExcelCellAddress;
import org.embulk.parser.poi_excel.visitor.embulk.CellVisitor;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnVisitor;
import org.embulk.spi.PageBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PoiExcelColumnVisitor implements ColumnVisitor {
    private final Logger log = LoggerFactory.getLogger(getClass());

    protected final PoiExcelVisitorValue visitorValue;
    protected final PageBuilder pageBuilder;
    protected final PoiExcelVisitorFactory factory;

    protected PoiExcelRecord record;

    public PoiExcelColumnVisitor(PoiExcelVisitorValue visitorValue) {
        this.visitorValue = visitorValue;
        this.pageBuilder = visitorValue.getPageBuilder();
        this.factory = visitorValue.getVisitorFactory();
    }

    public void setRecord(PoiExcelRecord record) {
        this.record = record;
    }

    @Override
    public final void booleanColumn(Column column) {
        visitCell0(column, factory.getBooleanCellVisitor());
    }

    @Override
    public final void longColumn(Column column) {
        visitCell0(column, factory.getLongCellVisitor());
    }

    @Override
    public final void doubleColumn(Column column) {
        visitCell0(column, factory.getDoubleCellVisitor());
    }

    @Override
    public final void stringColumn(Column column) {
        visitCell0(column, factory.getStringCellVisitor());
    }

    @Override
    public final void timestampColumn(Column column) {
        visitCell0(column, factory.getTimestampCellVisitor());
    }

    @Override
    public void jsonColumn(Column column) {
        throw new UnsupportedOperationException();
    }

    protected final void visitCell0(Column column, CellVisitor visitor) {
        if (log.isTraceEnabled()) {
            log.trace("{} start", column);
        }
        try {
            visitCell(column, visitor);
        } catch (Exception e) {
            String sheetName = visitorValue.getSheet().getSheetName();
            CellReference ref = record.getCellReference(visitorValue.getColumnBean(column));

            String message;
            if (ref != null) {
                message = MessageFormat.format("error at {0} cell={1}!{2}. {3}", column, sheetName, ref.formatAsString(), e.getMessage());
            } else {
                message = MessageFormat.format("error at {0} sheet={1}. {2}", column, sheetName, e.getMessage());
            }

            throw new RuntimeException(message, e);
        }
        if (log.isTraceEnabled()) {
            log.trace("{} end", column);
        }
    }

    protected void visitCell(Column column, CellVisitor visitor) {
        PoiExcelColumnBean bean = visitorValue.getColumnBean(column);
        PoiExcelColumnValueType valueType = bean.getValueType();
        PoiExcelCellAddress cellAddress = bean.getCellAddress();

        switch (valueType) {
        case FILE_NAME:
            visitor.visitFileName(column);
            return;
        case SHEET_NAME:
            if (cellAddress != null) {
                Sheet sheet = cellAddress.getSheet(record);
                visitor.visitSheetName(column, sheet);
            } else {
                visitor.visitSheetName(column);
            }
            return;
        case ROW_NUMBER:
            int rowIndex;
            if (cellAddress != null) {
                rowIndex = cellAddress.getRowIndex();
            } else {
                rowIndex = record.getRowIndex(bean);
            }
            visitor.visitRowNumber(column, rowIndex + 1);
            return;
        case COLUMN_NUMBER:
            int columnIndex;
            if (cellAddress != null) {
                columnIndex = cellAddress.getColumnIndex();
            } else {
                columnIndex = record.getColumnIndex(bean);
            }
            visitor.visitColumnNumber(column, columnIndex + 1);
            return;
        case CONSTANT:
            visitCellConstant(column, bean.getValueTypeSuffix(), visitor);
            return;
        default:
            break;
        }

        // assert valueType.useCell();
        Cell cell;
        if (cellAddress != null) {
            cell = cellAddress.getCell(record);
        } else {
            cell = record.getCell(bean);
        }
        if (cell == null) {
            visitCellNull(column);
            return;
        }
        switch (valueType) {
        case CELL_VALUE:
        case CELL_FORMULA:
            visitCellValue(bean, cell, visitor);
            return;
        case CELL_STYLE:
            visitCellStyle(bean, cell, visitor);
            return;
        case CELL_FONT:
            visitCellFont(bean, cell, visitor);
            return;
        case CELL_COMMENT:
            visitCellComment(bean, cell, visitor);
            return;
        case CELL_TYPE:
            visitCellType(bean, cell, cell.getCellType(), visitor);
            return;
        case CELL_CACHED_TYPE:
            if (cell.getCellType() == CellType.FORMULA) {
                visitCellType(bean, cell, cell.getCachedFormulaResultType(), visitor);
            } else {
                visitCellType(bean, cell, cell.getCellType(), visitor);
            }
            return;
        default:
            throw new UnsupportedOperationException(MessageFormat.format("unsupported value_type={0}", valueType));
        }
    }

    protected void visitCellConstant(Column column, String value, CellVisitor visitor) {
        if (value == null) {
            pageBuilder.setNull(column);
            return;
        }
        visitor.visitCellValueString(column, null, value);
    }

    protected void visitCellNull(Column column) {
        pageBuilder.setNull(column);
    }

    private void visitCellValue(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
        PoiExcelCellValueVisitor delegator = factory.getPoiExcelCellValueVisitor();
        delegator.visitCellValue(bean, cell, visitor);
    }

    private void visitCellStyle(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
        PoiExcelCellStyleVisitor delegator = factory.getPoiExcelCellStyleVisitor();
        delegator.visit(bean, cell, visitor);
    }

    private void visitCellFont(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
        PoiExcelCellFontVisitor delegator = factory.getPoiExcelCellFontVisitor();
        delegator.visit(bean, cell, visitor);
    }

    private void visitCellComment(PoiExcelColumnBean bean, Cell cell, CellVisitor visitor) {
        PoiExcelCellCommentVisitor delegator = factory.getPoiExcelCellCommentVisitor();
        delegator.visit(bean, cell, visitor);
    }

    private void visitCellType(PoiExcelColumnBean bean, Cell cell, CellType cellType, CellVisitor visitor) {
        PoiExcelCellTypeVisitor delegator = factory.getPoiExcelCellTypeVisitor();
        delegator.visit(bean, cell, cellType, visitor);
    }
}
