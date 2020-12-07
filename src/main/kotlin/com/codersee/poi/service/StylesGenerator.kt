package com.codersee.poi.service

import com.codersee.poi.model.CustomCellStyle
import org.apache.poi.ss.usermodel.*
import org.springframework.stereotype.Component

@Component
class StylesGenerator {

    fun prepareStyles(wb: Workbook): Map<CustomCellStyle, CellStyle> {
        val boldArial = createBoldArialFont(wb)
        val redBoldArial = createRedBoldArialFont(wb)

        val rightAlignedStyle = createRightAlignedStyle(wb)
        val greyCenteredBoldArialWithBorderStyle =
            createGreyCenteredBoldArialWithBorderStyle(wb, boldArial)
        val redBoldArialWithBorderStyle =
            createRedBoldArialWithBorderStyle(wb, redBoldArial)
        val rightAlignedDateFormatStyle =
            createRightAlignedDateFormatStyle(wb)

        return mapOf(
            CustomCellStyle.RIGHT_ALIGNED to rightAlignedStyle,
            CustomCellStyle.GREY_CENTERED_BOLD_ARIAL_WITH_BORDER to greyCenteredBoldArialWithBorderStyle,
            CustomCellStyle.RED_BOLD_ARIAL_WITH_BORDER to redBoldArialWithBorderStyle,
            CustomCellStyle.RIGHT_ALIGNED_DATE_FORMAT to rightAlignedDateFormatStyle
        )
    }

    private fun createBoldArialFont(wb: Workbook): Font {
        val font = wb.createFont()
        font.fontName = "Arial"
        font.bold = true
        return font
    }

    private fun createRedBoldArialFont(wb: Workbook): Font {
        val font = wb.createFont()
        font.fontName = "Arial"
        font.bold = true
        font.color = IndexedColors.RED.index
        return font
    }

    private fun createRightAlignedStyle(wb: Workbook): CellStyle {
        val style: CellStyle = wb.createCellStyle()
        style.alignment = HorizontalAlignment.RIGHT
        return style
    }

    private fun createGreyCenteredBoldArialWithBorderStyle(
        wb: Workbook,
        boldArial: Font
    ): CellStyle {
        val style = createBorderedStyle(wb)
        style.alignment = HorizontalAlignment.CENTER
        style.setFont(boldArial)
        style.fillForegroundColor = IndexedColors.GREY_25_PERCENT.getIndex();
        style.fillPattern = FillPatternType.SOLID_FOREGROUND;
        return style
    }

    private fun createRedBoldArialWithBorderStyle(
        wb: Workbook,
        redBoldArial: Font
    ): CellStyle {
        val style = createBorderedStyle(wb)
        style.setFont(redBoldArial)
        return style
    }

    private fun createRightAlignedDateFormatStyle(wb: Workbook): CellStyle {
        val style = wb.createCellStyle()
        style.alignment = HorizontalAlignment.RIGHT
        style.dataFormat = 14
        return style
    }

    private fun createBorderedStyle(wb: Workbook): CellStyle {
        val thin = BorderStyle.THIN
        val black = IndexedColors.BLACK.getIndex()
        val style = wb.createCellStyle()
        style.borderRight = thin
        style.rightBorderColor = black
        style.borderBottom = thin
        style.bottomBorderColor = black
        style.borderLeft = thin
        style.leftBorderColor = black
        style.borderTop = thin
        style.topBorderColor = black
        return style
    }
}
