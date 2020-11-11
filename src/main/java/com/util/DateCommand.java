package com.util;

import org.apache.poi.ss.usermodel.*;
import org.jxls.area.Area;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

import java.util.Objects;

public class DateCommand extends AbstractCommand {
    // 命令名
    public static final String COMMAND_NAME = "dateFormat";

    // 日期格式
    private String format;

    public void setFormat(String format) {
        this.format = format;
    }

    private Area area;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    /**
     * 在给定的单元格引用处应用命令
     * @param cellRef 必须在其中应用命令的单元格引用
     * @param context 要使用的Bean上下文
     * @return 转换后封闭命令区域的大小
     */
    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        // 没啥用，但得有
        Size size = area.applyAt(cellRef, context);

        // 允许XlsArea独立于任何特定实现而与Excel进行交互
        PoiTransformer transformer = (PoiTransformer) area.getTransformer();
        Workbook workbook = transformer.getWorkbook();
        // 读目标单元格（即自定义命令标注的单元格）
        Cell targetCell;
        {
            Sheet sheet = workbook.getSheet(cellRef.getSheetName());
            Row row = sheet.getRow(cellRef.getRow());
            targetCell = row.getCell(cellRef.getCol());
        }

        // 设置数据格式
        DataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat(this.format);
        targetCell.getCellStyle().setDataFormat(format);

        // 具体设置数据格式如下
//        CellStyle cellStyle = workbook.createCellStyle();
//        DataFormat dataformat= workbook.createDataFormat();
//        cellStyle.setDataFormat(dataformat.getFormat("yyyy-MM-dd"));
//        targetCell.setCellStyle(cellStyle);

        return size;
    }

    /**
     * 在此命令中添加区域，一定要有，否则报错
     * @param area 要添加的区域
     * @return 该命令实例
     */
    @Override
    public Command addArea(Area area) {
//        if (Objects.nonNull(this.area)) {
//            throw new IllegalArgumentException();
//        }
        this.area = area;
//        return super.addArea(area);
        return this;
    }
}
