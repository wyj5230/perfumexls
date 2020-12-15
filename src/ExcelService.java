import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelService
{
    public static Map<String, String> getSampleXls(String path) throws Exception {
        FileInputStream inputStream = new FileInputStream(new File(path));
        Map<String, String> sampleMap = new HashMap<>();
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet firstSheet = workbook.getSheetAt(1);
        int total = firstSheet.getLastRowNum();
        System.out.println("对照表共有" + total + "行");
        for (int i = 1; i <= total - 1; i++) {
            HSSFRow currentRow = firstSheet.getRow(i);
            sampleMap.put(currentRow.getCell(8).getStringCellValue(), currentRow.getCell(9).getStringCellValue());
        }
        System.out.println("对照表(去重)共有" + sampleMap.size() + "行");
        workbook.close();
        inputStream.close();
        return sampleMap;
    }

    public static List<Cloth> getActualXls(String path, int interval, Map<String, String> matchMap, List<Cloth> cloths) throws Exception {
        FileInputStream inputStream = new FileInputStream(new File(path));

        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet firstSheet = workbook.getSheetAt(0);
        int total = firstSheet.getLastRowNum();
        System.out.println("表"+path+"共有" + total + "行");
        if (interval>total -1){
            return null;
        }
        int hit = 0;
        for (int i = interval; i <= total - 1; i++) {
            HSSFRow currentRow = firstSheet.getRow(i);
            try {
                setCellTypeToString(currentRow, 2, 3, 5, 13);
            } catch (NullPointerException e) {
                continue;
            }
            if (currentRow.getCell(3) == null || currentRow.getCell(3).getStringCellValue().isEmpty()) {
                break;
            }
            if (currentRow.getCell(0).getCellType() != CellType.NUMERIC) {
                continue;
            }
            String cell3 = currentRow.getCell(3).getStringCellValue();
            if (cell3 != null) {
                String temp = matchMap.get(cell3);
                if (temp != null) {
                    cell3 = temp;
                    hit ++;
                }
            }

            Cloth cloth = new Cloth(currentRow.getCell(0).getNumericCellValue(),
                    currentRow.getCell(1).getStringCellValue(),
                    currentRow.getCell(2).getStringCellValue(),
                    cell3,
                    currentRow.getCell(4).getStringCellValue(),
                    currentRow.getCell(5).getStringCellValue(),
                    currentRow.getCell(6).getNumericCellValue(),
                    currentRow.getCell(7).getNumericCellValue(),
                    currentRow.getCell(8).getNumericCellValue(),
                    currentRow.getCell(9).getNumericCellValue(),
                    currentRow.getCell(10).getStringCellValue(),
                    currentRow.getCell(11).getStringCellValue(),
                    currentRow.getCell(12).getStringCellValue(),
                    currentRow.getCell(13).getStringCellValue());
            cloths.add(cloth);
        }
        System.out.println("共替换" + hit + "处");
        workbook.close();
        inputStream.close();
        return cloths;
    }

    private static void setCellTypeToString(HSSFRow currentRow, int i2, int i3, int i4, int i5)
    {
        currentRow.getCell(i2).setCellType(CellType.STRING);
        currentRow.getCell(i3).setCellType(CellType.STRING);
        currentRow.getCell(i4).setCellType(CellType.STRING);
        currentRow.getCell(i5).setCellType(CellType.STRING);
    }

    public static Title getTitle(String path) throws Exception {
        FileInputStream inputStream = new FileInputStream(new File(path));
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        HSSFSheet firstSheet = workbook.getSheetAt(0);
        HSSFRow currentRow = firstSheet.getRow(0);
        setCellTypeToString(currentRow, 0, 2, 3, 5);
        setCellTypeToString(currentRow, 6, 7, 8, 9);
        Title title = new Title(currentRow.getCell(0).getStringCellValue(),
                currentRow.getCell(1).getStringCellValue(),
                currentRow.getCell(2).getStringCellValue(),
                currentRow.getCell(3).getStringCellValue(),
                currentRow.getCell(4).getStringCellValue(),
                currentRow.getCell(5).getStringCellValue(),
                currentRow.getCell(6).getStringCellValue(),
                currentRow.getCell(7).getStringCellValue(),
                currentRow.getCell(8).getStringCellValue(),
                currentRow.getCell(9).getStringCellValue(),
                currentRow.getCell(10).getStringCellValue(),
                currentRow.getCell(11).getStringCellValue(),
                currentRow.getCell(12).getStringCellValue(),
                currentRow.getCell(13).getStringCellValue());
        workbook.close();
        inputStream.close();
        return title;
    }

    public static void printMergedXls(String fileName, Title title, List<Cloth> cloths) throws IOException{
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("Sheet1");
        createRowFromClothList(title, cloths, sheet1, workbook);
        FileOutputStream out =
                new FileOutputStream(new File("./" + fileName + ".xls"));
        workbook.write(out);
        out.close();
        System.out.println("成功生成新表：" + fileName + ".xls");
        workbook.close();
    }
    private static void createRowFromClothList(Title title, List<Cloth> clothList, Sheet sheet1, Workbook workbook) {
        createTitle(sheet1, title, workbook);
        for (int i = 1; i <= clothList.size(); i++) {
            Cloth cloth = clothList.get(i - 1);
            createRowFromCloth(sheet1, i, cloth, workbook);
        }
    }

    private static void createTitle(Sheet sheet1, Title title, Workbook workbook) {
        Row currentRow = sheet1.createRow(0);
        Cell cell0 = currentRow.createCell(0);
        cell0.setCellValue(title.getDate());
        Cell cell1 = currentRow.createCell(1);
        cell1.setCellValue(title.getStore());
        Cell cell2 = currentRow.createCell(2);
        cell2.setCellValue(title.getStoreCode());
        Cell cell3 = currentRow.createCell(3);
        cell3.setCellValue(title.getBarCode());
        Cell cell4 = currentRow.createCell(4);
        cell4.setCellValue(title.getSaleType());
        Cell cell5 = currentRow.createCell(5);
        cell5.setCellValue(title.getQty());
        Cell cell6 = currentRow.createCell(6);
        cell6.setCellValue(title.getNetAmount1());
        Cell cell7 = currentRow.createCell(7);
        cell7.setCellValue(title.getNetAmount2());
        Cell cell8 = currentRow.createCell(8);
        cell8.setCellValue(title.getGrossAmount());
        Cell cell9 = currentRow.createCell(9);
        cell9.setCellValue(title.getUnitListPrice());
        Cell cell10 = currentRow.createCell(10);
        cell10.setCellValue(title.getStyle());
        Cell cell11 = currentRow.createCell(11);
        cell11.setCellValue(title.getFabric());
        Cell cell12 = currentRow.createCell(12);
        cell12.setCellValue(title.getColor());
        Cell cell13 = currentRow.createCell(13);
        cell13.setCellValue(title.getSize());
    }

    private static void createRowFromCloth(Sheet sheet1, int i, Cloth cloth, Workbook workbook) {
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        CellStyle dateStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("yyyy/mm/dd"));

        Row currentRow = sheet1.createRow(i);
        Cell cell0 = currentRow.createCell(0);
        cell0.setCellStyle(dateStyle);
        cell0.setCellValue(cloth.getDate());
        Cell cell1 = currentRow.createCell(1);
        cell1.setCellValue(cloth.getStore());
        Cell cell2 = currentRow.createCell(2);
        cell2.setCellValue(cloth.getStoreCode());
        Cell cell3 = currentRow.createCell(3);
        cell3.setCellValue(cloth.getBarCode());
        Cell cell4 = currentRow.createCell(4);
        cell4.setCellValue(cloth.getSaleType());
        Cell cell5 = currentRow.createCell(5);
        cell5.setCellValue(cloth.getQty());
        Cell cell6 = currentRow.createCell(6);
        cell6.setCellStyle(dataStyle);
        cell6.setCellValue(cloth.getNetAmount1());
        Cell cell7 = currentRow.createCell(7);
        cell7.setCellStyle(dataStyle);
        cell7.setCellValue(cloth.getNetAmount2());
        Cell cell8 = currentRow.createCell(8);
        cell8.setCellStyle(dataStyle);
        cell8.setCellValue(cloth.getGrossAmount());
        Cell cell9 = currentRow.createCell(9);
        cell9.setCellStyle(dataStyle);
        cell9.setCellValue(cloth.getUnitListPrice());
        Cell cell10 = currentRow.createCell(10);
        cell10.setCellValue(cloth.getStyle());
        Cell cell11 = currentRow.createCell(11);
        cell11.setCellValue(cloth.getFabric());
        Cell cell12 = currentRow.createCell(12);
        cell12.setCellValue(cloth.getColor());
        Cell cell13 = currentRow.createCell(13);
        cell13.setCellValue(cloth.getSize());
    }

    public static void copySheets(HSSFWorkbook newWorkbook, HSSFSheet newSheet, HSSFSheet sheet, boolean copyStyle){
        int newRownumber = newSheet.getLastRowNum();
        int maxColumnNum = 0;
        Map<Integer, HSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, HSSFCellStyle>() : null;

        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            HSSFRow srcRow = sheet.getRow(i);
            HSSFRow destRow = newSheet.createRow(i + newRownumber);
            if (srcRow != null) {
                copyRow(newWorkbook, sheet, newSheet, srcRow, destRow, styleMap);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }

    public static void copyRow(HSSFWorkbook newWorkbook, HSSFSheet srcSheet, HSSFSheet destSheet, HSSFRow srcRow, HSSFRow destRow, Map<Integer, HSSFCellStyle> styleMap) {
        destRow.setHeight(srcRow.getHeight());
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            HSSFCell oldCell = srcRow.getCell(j);
            HSSFCell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                copyCell(newWorkbook, oldCell, newCell, styleMap);
            }
        }
    }

    public static void copyCell(HSSFWorkbook newWorkbook, HSSFCell oldCell, HSSFCell newCell, Map<Integer, HSSFCellStyle> styleMap) {
        if (oldCell == null) {
            return;
        }
        if(styleMap != null) {
            int stHashCode = oldCell.getCellStyle().hashCode();
            HSSFCellStyle newCellStyle = styleMap.get(stHashCode);
            if(newCellStyle == null){
                newCellStyle = newWorkbook.createCellStyle();
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                styleMap.put(stHashCode, newCellStyle);
            }
            newCell.setCellStyle(newCellStyle);
        }
        switch(oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getRichStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }
    }
}
