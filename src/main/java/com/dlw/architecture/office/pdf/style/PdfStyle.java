package com.dlw.architecture.office.pdf.style;

import com.dlw.architecture.office.enums.ColorType;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.io.OutputStream;

/**
 * @author dengliwen
 * @date 2020/6/18
 * @desc 默认PDF样式
 * @since 4.0.0
 */
@Slf4j
public class PdfStyle {

    private PdfStyle(){}

    private static BaseFont bfChinese;
    static {
        try {
            bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
                    BaseFont.NOT_EMBEDDED);
        } catch (DocumentException | IOException e) {
            log.error(e.getMessage(), e);
        }
    }

    /**
     * 获取文档 默认值使用
     * @return
     */
    public static Document getDocument() {
        return getDocument(PageSize.A4);
    }

    public static Document getDocument(Rectangle pageSize) {
        return new Document(pageSize);
    }

    private static BaseColor getColor(ColorType colorType) {
        BaseColor color;
        switch (colorType) {
            case RED:
                color = BaseColor.RED;
                break;
            case BLUE:
                color = BaseColor.BLUE;
                break;
            case CYAN:
                color = BaseColor.CYAN;
                break;
            case GRAY:
                color = BaseColor.GRAY;
                break;
            case PINK:
                color = BaseColor.PINK;
                break;
            case BLACK:
                color = BaseColor.BLACK;
                break;
            case GREEN:
                color = BaseColor.GREEN;
                break;
            case WHITE:
                color = BaseColor.WHITE;
                break;
            case ORANGE:
                color = BaseColor.ORANGE;
                break;
            case YELLOW:
                color = BaseColor.YELLOW;
                break;
            case MAGENTA:
                color = BaseColor.MAGENTA;
                break;
            case DARK_GRAY:
                color = BaseColor.DARK_GRAY;
                break;
            case LIGHT_GRAY:
                color = BaseColor.LIGHT_GRAY;
                break;
            default:
                color = BaseColor.BLACK;
                break;
        }
        return color;
    }


    /**
     * 获取默认样式 支持中文
     * @return
     */
    public static Font getTextFont(int size, ColorType colorType) {
        return new Font(bfChinese,size,Font.NORMAL,getColor(colorType));
    }

    public static Font getBoldFont(int size, ColorType colorType) {
        return new Font(bfChinese,size,Font.BOLD,getColor(colorType));
    }

    public static Font getFirstTitleFont(int size, ColorType colorType) {
        return new Font(bfChinese,size,Font.BOLD,getColor(colorType));
    }

    public static Font getSecondTitleFont(int size, ColorType colorType) {
        return new Font(bfChinese,size,Font.BOLD,getColor(colorType));
    }

    public static Font getUnderlineFont(int size, ColorType colorType) {
        return new Font(bfChinese,size,Font.UNDERLINE,getColor(colorType));
    }
}
