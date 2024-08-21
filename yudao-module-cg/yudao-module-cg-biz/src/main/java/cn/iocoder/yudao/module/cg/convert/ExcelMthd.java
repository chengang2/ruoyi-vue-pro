package cn.iocoder.yudao.module.cg.convert;

import cn.iocoder.yudao.module.cg.controller.admin.excel.vo.FYDataModel;
import com.google.zxing.BarcodeFormat;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import io.netty.util.internal.StringUtil;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.*;

@Component
public class ExcelMthd {

    public static int lastPageFirstRow;
    //填充数据到excel
    public void reportStaticExcel(String excelPath, int sheetIndex, Map<String, String> dataMap) {
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(excelPath))) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        boolean isReplaced = false; // 标记是否有替换
                        for (Map.Entry<String, String> entry : dataMap.entrySet()) {
                            String placeholder = entry.getKey();
                            String replacement = entry.getValue();
                            // 如果当前单元格值包含占位符
                            if (cellValue.contains(placeholder)) {
                                cellValue = cellValue.replace(placeholder, replacement);
                                isReplaced = true; // 标记替换成功
                            }
                        }
                        if (isReplaced) {
                            cell.setCellValue(cellValue); // 只有在发生替换时才更新单元格值
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(excelPath)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    //添加图片到excel，比如封页
    //imagePoints:[图片插入位置左边距,图片插入位置上边距,图片宽度,图片高度] 如{{15,15,200,150}},宽高为-1时,将使用原始大小
    public  boolean addImagesToExcel(String workbookPath, int sheetIndex, String[] imagePaths, float[][] imagePoints) {
        boolean flag = true;
        FileInputStream fileInputStream = null;
        FileOutputStream fileOutputStream = null;
        Workbook workbook = null;

        try {
            fileInputStream = new FileInputStream(workbookPath);
            workbook = WorkbookFactory.create(fileInputStream);

            Sheet sheet = workbook.getSheetAt(sheetIndex);

            // Drawing patriarch to create anchors
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            for (int i = 0; i < imagePoints.length; i++) {
                // Load the image file
                InputStream inputStream = new FileInputStream(imagePaths[i]);
                byte[] bytes = IOUtils.toByteArray(inputStream);
                inputStream.close();

                // Add picture data to the workbook
                int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

                // Create an anchor point
                CreationHelper helper = workbook.getCreationHelper();
                ClientAnchor anchor = helper.createClientAnchor();

                // Set top-left corner for the image
//                anchor.setCol1((int) imagePoints[i][0]);
//                anchor.setRow1((int) imagePoints[i][1]);
//                // Calculate bottom-right corner based on the width and height in points
//                // Assuming the width and height provided are in pixels, we convert them to points
//                double widthInPoints = imagePoints[i][2] * Units.PIXEL_DPI / Units.DEFAULT_CHARACTER_WIDTH;
//                double heightInPoints = imagePoints[i][3] * Units.PIXEL_DPI / Units.DEFAULT_CHARACTER_WIDTH;
//                anchor.setCol2(anchor.getCol1() + (int) Math.round(widthInPoints));
//                anchor.setRow2(anchor.getRow1() + (int) Math.round(heightInPoints));
                float pLeft = imagePoints[i][0];
                float pTop = imagePoints[i][1];
                float fWidth = imagePoints[i][2];
                float fHeight = imagePoints[i][3];
                // 使用列和行作为锚点设置左上角
                // 假设pLeft和pTop以点为单位，将它们转换为Excel的单位
                int col1 = Math.round(pLeft / Units.DEFAULT_CHARACTER_WIDTH);
                int row1 = Math.round(pTop / Units.DEFAULT_CHARACTER_WIDTH);  // We use the same approximation for height units
                anchor.setCol1((int) pLeft);
                anchor.setRow1((int) pTop);
                // dx1和dy1是相对于左上角单元格的偏移量
                // 如果位置已由col1和row1定义，则应将它们设置为零
                anchor.setDx1(0);
                anchor.setDy1(0);
                // 宽度和高度（fWidth和fHeight）用于计算图像的右下角
                // 将这些尺寸从点转换为Excel的单位（EMU - 英制度量单位）
                // 右下角通过将宽度和高度添加到左上角来确定
                anchor.setDx2((int) (fWidth * Units.EMU_PER_POINT));
                anchor.setDy2((int) (fHeight * Units.EMU_PER_POINT));
                // Creates the image
                Picture pict = drawing.createPicture(anchor, pictureIdx);
                // Reset image to the original size
                pict.resize();
            }

            // Write the output to a file
            fileOutputStream = new FileOutputStream(workbookPath);
            workbook.write(fileOutputStream);

        } catch (Exception e) {
            e.printStackTrace();
            flag = false;
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
                if (workbook != null) {
                    workbook.close();
                }
                if (fileInputStream != null) {
                    fileInputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return flag;
    }

    // 修改excel sheet的名称
    public  void sheetRename(String workbookPath, int sheetIndex, String newSheetName) {
        FileInputStream fileInputStream = null;
        FileOutputStream fileOutputStream = null;
        Workbook workbook = null;

        try {
            fileInputStream = new FileInputStream(workbookPath);
            workbook = WorkbookFactory.create(fileInputStream);

            if (workbook.getNumberOfSheets() <= sheetIndex) {
                throw new IllegalArgumentException("Sheet index is out of bounds.");
            }

            Sheet sheet = workbook.getSheetAt(sheetIndex);
            if (newSheetName != null && !newSheetName.trim().isEmpty()) {
                workbook.setSheetName(sheetIndex, newSheetName);
            }

            // Write the changes to the workbook
            fileOutputStream = new FileOutputStream(workbookPath);
            workbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
            // You might want to implement a logging mechanism instead of printing the stack trace.
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
                if (workbook != null) {
                    workbook.close();
                }
                if (fileInputStream != null) {
                    fileInputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public  void protectWorksheet(String filePath, int sheetIndex, String password ) {

        try {
            Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            sheet.protectSheet(password);
            // Apache POI does not directly support all the granular permissions shown in your C# example
            // You would typically handle these permissions using the protectSheet method and setting permissions
            // on the workbook's windows and workbook structure if needed.

            // Save changes to the workbook here if needed

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public  void unprotectWorkSheet(Workbook wb, int sheetIndex, String password) {
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            // Apache POI does not have a direct method to unprotect a sheet with a password
            // You will need to manipulate the underlying XML structure to remove the protection
            // Here is an example using XSSF sheet (for .xlsx files)
//            这个判断的目的是确保我们正在处理的是一个 XSSFSheet 对象，因为 XSSFSheet 特定于处理 .xlsx 文件格式（Office Open XML 格式）。
//            XSSFSheet 提供了对底层 XML 结构的访问，我们可以通过它来操作工作表的保护设置。
//            在 Apache POI 库中，有两种主要的工作表实现：
//            HSSFSheet：用于处理 .xls 文件（Excel 97-2003 格式）。
//            XSSFSheet：用于处理 .xlsx 文件（Excel 2007 及更高版本的格式）。
//            由于解除保护的操作涉及到底层 XML 结构的操作，只有 XSSFSheet 提供了相关的 API。
            if (sheet instanceof XSSFSheet) {
                XSSFSheet xssfSheet = (XSSFSheet) sheet;
                CTWorksheet ctWorksheet = xssfSheet.getCTWorksheet();
                if (ctWorksheet.isSetSheetProtection()) {
                    ctWorksheet.unsetSheetProtection();
                }
            } else if (sheet instanceof HSSFSheet) {
                // 处理 .xls 文件
                HSSFSheet hssfSheet = (HSSFSheet) sheet;
                // 解除保护
                hssfSheet.protectSheet(null); // 设置密码为空来解除保护
            }

        } catch (Exception ex) {
            // Handle exception
            ex.printStackTrace();
        }
    }


    // excel插入另一个sheet
    public void sheetInsert(Workbook fromWb, int fromIndex, Workbook toWb, int toIndex, String sType, String newSheetName,String toWbPath) {
        try {
            if (fromWb.getNumberOfSheets() < fromIndex) return;
            if (toWb.getNumberOfSheets() < toIndex) return;

            Sheet fromSheet = fromWb.getSheetAt(fromIndex); // 获取要复制的sheet页，索引从0开始
            // 创建一个新的工作表
            Sheet newSheet = toWb.createSheet(newSheetName);
            // 复制工作表内容
            copySheet(fromSheet, newSheet, fromWb, toWb);
            // 设置新Sheet的顺序
            if (sType.equals("BEFORE") && toIndex > 0) {
                toWb.setSheetOrder(newSheetName, toIndex);
            }
            if (sType.equals("AFTER")) {
                toWb.setSheetOrder(newSheetName, toIndex + 1);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }finally {
             //保存目标工作簿
            try (FileOutputStream outputStream = new FileOutputStream(toWbPath)) {
                toWb.write(outputStream);
            }catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private void copySheet(Sheet sourceSheet, Sheet targetSheet, Workbook sourceWorkbook, Workbook targetWorkbook) {
        // 复制行
        for (Row sourceRow : sourceSheet) {
            Row targetRow = targetSheet.createRow(sourceRow.getRowNum());
            copyRow(sourceRow, targetRow, sourceWorkbook, targetWorkbook);
        }

        // 复制列样式
        for (int i = 0; i< sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress sourceRegion = sourceSheet.getMergedRegion(i);
            CellRangeAddress targetRegion = new CellRangeAddress(sourceRegion.getFirstRow(), sourceRegion.getLastRow(),
                    sourceRegion.getFirstColumn(), sourceRegion.getLastColumn());
            targetSheet.addMergedRegion(targetRegion);
        }
    }

    private  void copyRow(Row sourceRow, Row targetRow, Workbook sourceWorkbook, Workbook targetWorkbook) {
        // 复制单元格
        for (Cell sourceCell : sourceRow) {
            Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex());
            copyCell(sourceCell, targetCell, sourceWorkbook, targetWorkbook);
        }

        // 复制行样式
        targetRow.setHeight(sourceRow.getHeight());
    }

    private  void copyCell(Cell sourceCell, Cell targetCell, Workbook sourceWorkbook, Workbook targetWorkbook) {
        // 复制单元格样式
        // 创建一个新的CellStyle对象
        CellStyle newCellStyle = targetWorkbook.createCellStyle();
        // 获取旧的CellStyle对象
        CellStyle oldCellStyle = sourceCell.getCellStyle();

        // 复制样式
        newCellStyle.cloneStyleFrom(oldCellStyle); // 在POI 3.15中引入的方法

        targetCell.setCellStyle(newCellStyle);

        // 复制单元格值
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setCellValue("");
                break;
            default:
                break;
        }
    }

    public  String reportFy(String workbookPath,
                                  int sheetIndex,
                                  FYDataModel fyDataModel,
                                  String[] columnsWidth) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(new FileInputStream(workbookPath));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        List<Map<String, String>> fyDataTable = fyDataModel.getFYDataTable();
        List<String> fyImageList = fyDataModel.getFYImageArray();
        List<String> fyColumnListToMerge = fyDataModel.getFYColumnListToMerge();
        String[][] dArray =convertToTwoDimensionalArray(fyDataTable);

        // fyColumnListToMerge 转为数组
        String[] colListC = fyColumnListToMerge.toArray(new String[0]);
        // 判断fyImageArray是否为null
        String[] fyImageArray = null;
        if (fyImageList != null) {
            fyImageArray = fyImageList.toArray(new String[0]);
        }

        String testType = "测试分类";
        String testType2 = "测试分类2";

        // 这个方法的目的是为了检验 传递过来的附页数据源中有没有 一列叫做 “测试分类”
        boolean cln = CheckColumnName(dArray, testType);
        boolean cln2 = CheckColumnName(dArray, testType2);

        Sheet sheet = wb.getSheetAt(sheetIndex);
        // 获取sheet活动区域的最大行列号
        Point pSheet = getSheetMaxRowCol(sheet);
        int maxRowIndex = pSheet.x -1;
        int maxColIndex = pSheet.y -1 ;
        //获取合并列列名在模板中的顺序
        int[] colseq = getArraySequen(sheet, colListC, "&[", "]", maxRowIndex, maxColIndex);

        //2:填充表格数据
        int[] rowRange = reportTable(wb,sheetIndex,dArray,maxColIndex);
        int startRow = rowRange[0];
        int endRow = rowRange[1];
        //3:设置附页列宽
        if (columnsWidth != null && columnsWidth.length > 0)//columnsWidth,如果列宽数组有效,则设置附页列宽
        {
            setSheetColumnsWidth(wb, sheetIndex, columnsWidth);
        }
        //4:合并单元格
        if (cln) {
            // 行合并测试分类，检验检测项目，分析项目
            mergeRowTestAndAnalyte2(wb, sheetIndex, startRow, endRow, maxColIndex, testType); // 合并检测项目和分析项
            //mergeRows(wb, sheetIndex, startRow, endRow,unpivotRange,maxColIndex, unpivotMerge); // 合并转置列
            mergeCells3(wb, sheetIndex, colseq, startRow, endRow); // 合并相同检测项目的同列数据

            //wb = setAutoRowHeight(wb, sheetIndex, startRow, 0, maxColIndex); // 设置行高
            dealMergedAreaInPages_new(wb, sheetIndex, startRow,maxColIndex, colseq); // 跨页的要分开
            mergeTestCell2(wb, sheetIndex, testType, "检验检测项目", "分析项", maxColIndex); // 合并表头的"检测项目","分析项"2个单元格
        } else if (cln2) { // 附页模板中第二列不再是检验检测测试项目而是测试分类2
            // 行合并测试分类，检验检测项目，分析项目
            mergeRowTestAndAnalyte22(wb, sheetIndex, startRow, endRow, maxColIndex, testType2); // 合并检测项目和分析项, 测试分类2
            //mergeRows(wb, sheetIndex, startRow,endRow, unpivotRange, maxColIndex, unpivotMerge); // 合并转置列
            mergeCells33(wb, sheetIndex, colseq, startRow, endRow); // 合并相同检测项目的同列数据

            //wb = setAutoRowHeight(wb, sheetIndex, startRow, 0, maxColIndex); // 设置行高
            dealMergedAreaInPages_new(wb, sheetIndex, startRow,maxColIndex, colseq); // 跨页的要分开
            mergeTestCell22(wb, sheetIndex, testType2, "检验检测项目", "分析项", maxColIndex); // 合并表头的"检测项目","分析项"2个单元格
        } else {
            mergeRowTestAndAnalyte(wb, sheetIndex, startRow, endRow, maxColIndex); // 合并检测项目和分析项
            //mergeRows(wb, sheetIndex, startRow, endRow,unpivotRange, maxColIndex, unpivotMerge); // 合并转置列
            mergeCells1(wb, sheetIndex, colseq, startRow, endRow); // 合并相同检测项目的同列数据

            //wb = setAutoRowHeight(wb, sheetIndex, startRow, 0, maxColIndex); // 设置行高
            dealMergedAreaInPages_new(wb, sheetIndex, startRow,maxColIndex, colseq); // 跨页的要分开
            mergeTestCell(wb, sheetIndex, "检验检测项目", "分析项", maxColIndex); // 合并表头的"检测项目","分析项"2个单元格
        }

        //5:图片插入
        if (fyImageArray != null) {
            //todo: 图片插入
            //addImagesToRowWithAdjustedHeight(workbookPath, 0, fyImageArray, endRow+1);
        }
        //6: 保存为pdf
//        saveWorkBookAsPdf(wb, printDir);
        try (FileOutputStream fos = new FileOutputStream(workbookPath)) {
            wb.write(fos);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return null;
    }
    private int[] reportTable(Workbook wb, int sheetIndex, String[][] dArray, int maxColIndex){
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);
            int maxRowIndex = getSheetMaxRowCol(sheet).x - 1 ;
            String[] tableHead = dArray[0]; // The first row is the table header

            int[] colHeadSeq = getArraySequen(sheet, tableHead, "&[", "]", maxRowIndex, maxColIndex); // Get the sequence of the table header in the template

            String[][] seqArray2 = getSequenArray22(colHeadSeq, dArray);

            List<Integer> colHeadSeqList = new ArrayList<>();
            for (int seq : colHeadSeq) {
                colHeadSeqList.add(seq);
            }
            int index = colHeadSeqList.indexOf(0);
            String cellPosiValue = tableHead[index];

            Point startPosition = selectPosition(sheet, "&[" + cellPosiValue + "]", maxRowIndex, maxColIndex);
            int row = startPosition.x;
            int col = startPosition.y;

            if (row == -1 || col == -1) return new int[]{0, 0};
            // 先插入行,再填充数据, seqArray2.length为矩形数组的行数
            for (int i = 1; i < seqArray2.length; i++) { // 少插入一行是因为二维数组有表头信息
                sheet.shiftRows(row+1, sheet.getLastRowNum(), 1);

                // 复制行格式
                Row oldRow = sheet.getRow(row);
                Row newRow = sheet.createRow(row + 1);
                newRow.setHeight(oldRow.getHeight()); // 复制行高
//                if (oldRow != null) {
//                    copyRowStyle(oldRow, newRow);
//                }
            }
            // 直接填充数组
            for (int i = 1; i < seqArray2.length; i++) {
                Row oldRow = sheet.getRow(row);
                Row currentRow = sheet.getRow(row + i);
                if (currentRow == null) currentRow = sheet.createRow(row + i);
                for (int j = 0; j < seqArray2[i].length; j++) {
                    Cell oldCell = oldRow.getCell(j);
                    Cell cell = currentRow.createCell(j);
                    cell.setCellValue(seqArray2[i][j]);
                    copyCellStyle(oldCell, cell);
                    //copyCellStyle(sheet, row, col, currentRow.getRowNum(), cell.getColumnIndex());

                }
            }
            sheet.shiftRows(row+1, sheet.getLastRowNum(), -1);

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();
            maxRowIndex = row + seqArray2.length - 1;

            //wb.write(new FileOutputStream("/Users/chengang/Downloads/26.xls"));
            return new int[] { row, maxRowIndex };
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
            return new int[] { 0, 0 };
        }

    }
    private Point getSheetMaxRowCol(Sheet sheet) {
        try {
            // Apache POI's Sheet.getLastRowNum() returns the last row number.
            int lastRowNum = sheet.getLastRowNum();

            // Finding the last column index requires iterating over the rows and checking for the last cell number.
            int lastColNum = 0;
            for (int rowNum = 0; rowNum <= lastRowNum; rowNum++) {
                if (sheet.getRow(rowNum) != null && sheet.getRow(rowNum).getLastCellNum() > lastColNum) {
                    lastColNum = sheet.getRow(rowNum).getLastCellNum();
                }
            }

            // Apache POI's cells are 0-based, so we add 1 to get the actual count.
            int endRow = lastRowNum + 1;
            int endCol = lastColNum; // getLastCellNum() already returns the number of columns as 1-based.

            if (endRow > 0 && endCol > 0) {
                return new Point(endRow, endCol);
            } else {
                return new Point(2000, 1000);
            }
        } catch (Exception e) {
            return new Point(2000, 1000);
        }
    }

    public String[][] convertToTwoDimensionalArray(List<Map<String, String>> dataTable) {
        if (dataTable == null || dataTable.isEmpty()) {
            return new String[0][0];
        }

        // 确保 "序号" 列存在
        if (!dataTable.get(0).containsKey("序号")) {
            throw new IllegalArgumentException("数据表中必须包含 '序号' 列");
        }

        // 按 "序号" 排序
        dataTable.sort(Comparator.comparing(row -> Integer.parseInt(row.get("序号"))));

        // 获取表头（第一维）
        Set<String> headers = dataTable.get(0).keySet();
        List<String> headerList = new ArrayList<>(headers);

        // 确保 "序号" 是第一列
        headerList.remove("序号");
        headerList.add(0, "序号");

        String[] headerArray = headerList.toArray(new String[0]);

        // 创建二维数组
        String[][] result = new String[dataTable.size() + 1][headerArray.length];

        // 填充表头
        result[0] = headerArray;

        // 填充数据
        for (int i = 0; i < dataTable.size(); i++) {
            Map<String, String> row = dataTable.get(i);
            for (int j = 0; j < headerArray.length; j++) {
                result[i + 1][j] = row.get(headerArray[j]);
            }
        }
        return result;
    }
    private void setSheetColumnsWidth(Workbook wb, int sheetIndex, Object[] columnsWidth) {
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            // Assuming columnsWidth is an array of numbers (e.g., Double or Integer)
            for (int i = 0; i < columnsWidth.length; i++) {
                // Apache POI sets column width in units of 1/256th of a character width
                // Excel's width unit is character width, this conversion might be needed.
                int width = (int) (256 * (double) columnsWidth[i]);
                sheet.setColumnWidth(i, width);

                // Log the column index and width if needed
                // System.out.println("Current column: " + (i + 1) + ", Column width: " + columnsWidth[i]);
            }
        } catch (Exception ex) {
            // Log the exception as needed
            // This is a placeholder for logging the exception
            ex.printStackTrace();
        }
    }
    private boolean CheckColumnName(String[][] dArray, String dataColumnName) {
        // 将Object数组的第一行转换为String数组
        String[] tableHead = dArray[0];

        // 检查tableHead数组中是否包含dataColumnName
        if (tableHead != null) {
            for (String columnName : tableHead) {
                if (columnName.equals(dataColumnName)) {
                    return true;
                }
            }
        }

        return false;
    }

    public int[] getArraySequen(Sheet sheet, String[] values, String prefix, String suffix, int endRow, int endCol) {
        int[] sequen = new int[values.length];

        for (int i = 0; i < values.length; i++) {
            Point position = selectPosition(sheet, prefix + values[i] + suffix, endRow, endCol);
            sequen[i] = position.y;
        }

        return sequen;
    }
    public Point selectPosition(Sheet sheet, String value, int endRow, int endCol) {
        Point p = new Point(-1, -1);
        try {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cellMatchesValue(cell, value)) {
                        p.x = row.getRowNum();
                        p.y = cell.getColumnIndex();
                        return p;
                    }
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return p;
    }
    private boolean cellMatchesValue(Cell cell, String value) {
        if (cell == null) return false;

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().toLowerCase().contains(value.toLowerCase());
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue()).contains(value);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).contains(value);
            case FORMULA:
                try {
                    return cell.getStringCellValue().toLowerCase().contains(value.toLowerCase());
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue()).contains(value);
                    } catch (Exception ex) {
                        return false;
                    }
                }
            default:
                return false;
        }
    }

    public String[][] getSequenArray22(int[] sequen, String[][] array) {
        try {
            int rowCount = array.length;
            int colCount = array[0].length;
            String[][] sequenArray2 = new String[rowCount][colCount];

            for (int i = 0; i < rowCount; i++) {
                for (int j = 0; j < colCount; j++) {
                    // Note: Array indices in Java are zero-based, unlike C# where arrays can be one-based.
                    // Therefore, we subtract 1 from sequen[j] to convert to zero-based index.
                    sequenArray2[i][sequen[j]] = array[i][j];
                }
            }
            return sequenArray2;
        } catch (Exception ex) {
            // Replace with your logging mechanism.
            ex.printStackTrace();
            // In case of exception, return the original array. You may want to handle this differently.
            return array;
        }
    }
    private void copyCellStyle(Cell oldCell, Cell newCell) {
        newCell.setCellStyle(oldCell.getCellStyle());
//        Workbook workbook = oldCell.getSheet().getWorkbook();
//        CellStyle newCellStyle = workbook.createCellStyle();
//        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//        newCell.setCellStyle(newCellStyle);
//
//        // 复制单元格类型和值
//        switch (oldCell.getCellType()) {
//            case STRING:
//                newCell.setCellValue(oldCell.getStringCellValue());
//                break;
//            case NUMERIC:
//                if (DateUtil.isCellDateFormatted(oldCell)) {
//                    newCell.setCellValue(oldCell.getDateCellValue());
//                } else {
//                    newCell.setCellValue(oldCell.getNumericCellValue());
//                }
//                break;
//            case BOOLEAN:
//                newCell.setCellValue(oldCell.getBooleanCellValue());
//                break;
//            case FORMULA:
//                newCell.setCellFormula(oldCell.getCellFormula());
//                break;
//            case BLANK:
//                newCell.setBlank();
//                break;
//            case ERROR:
//                newCell.setCellErrorValue(oldCell.getErrorCellValue());
//                break;
//            default:
//                break;
//        }
//        // 复制单元格的其他属性
//        copyCellFont(oldCell, newCell);
//        copyCellBorders(oldCell, newCell);
//        copyCellAlignment(oldCell, newCell);
    }
    private void copyCellFont(Cell oldCell, Cell newCell) {
        Workbook workbook = oldCell.getSheet().getWorkbook();
        Font oldFont = workbook.getFontAt(oldCell.getCellStyle().getFontIndex());
        Font newFont = workbook.createFont();
        newFont.setBold(oldFont.getBold());
        newFont.setColor(oldFont.getColor());
        newFont.setFontHeight(oldFont.getFontHeight());
        newFont.setFontName(oldFont.getFontName());
        newFont.setItalic(oldFont.getItalic());
        newFont.setStrikeout(oldFont.getStrikeout());
        newFont.setTypeOffset(oldFont.getTypeOffset());
        newFont.setUnderline(oldFont.getUnderline());
        newFont.setCharSet(oldFont.getCharSet());

        CellStyle newCellStyle = newCell.getCellStyle();
        newCellStyle.setFont(newFont);
        newCell.setCellStyle(newCellStyle);
    }

    private void copyCellBorders(Cell oldCell, Cell newCell) {
        CellStyle newCellStyle = newCell.getCellStyle();
        newCellStyle.setBorderTop(oldCell.getCellStyle().getBorderTop());
        newCellStyle.setBorderBottom(oldCell.getCellStyle().getBorderBottom());
        newCellStyle.setBorderLeft(oldCell.getCellStyle().getBorderLeft());
        newCellStyle.setBorderRight(oldCell.getCellStyle().getBorderRight());
        newCellStyle.setTopBorderColor(oldCell.getCellStyle().getTopBorderColor());
        newCellStyle.setBottomBorderColor(oldCell.getCellStyle().getBottomBorderColor());
        newCellStyle.setLeftBorderColor(oldCell.getCellStyle().getLeftBorderColor());
        newCellStyle.setRightBorderColor(oldCell.getCellStyle().getRightBorderColor());
        newCell.setCellStyle(newCellStyle);
    }

    private void copyCellAlignment(Cell oldCell, Cell newCell) {
        CellStyle newCellStyle = newCell.getCellStyle();
        newCellStyle.setAlignment(oldCell.getCellStyle().getAlignment());
        newCellStyle.setVerticalAlignment(oldCell.getCellStyle().getVerticalAlignment());
        newCellStyle.setWrapText(oldCell.getCellStyle().getWrapText());
        newCellStyle.setRotation(oldCell.getCellStyle().getRotation());
        newCellStyle.setIndention(oldCell.getCellStyle().getIndention());
        newCell.setCellStyle(newCellStyle);
    }
    private void mergeRowTestAndAnalyte(Workbook wb, int sheetIndex, int startRow, int endRow,int maxCol) {
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            Point startColPoint = selectPosition(sheet, "检验检测项目", endRow, maxCol);
            Point endColPoint = selectPosition(sheet, "分析项", endRow, maxCol);

            int startCol = startColPoint.y;
            int endCol = endColPoint.y;

            if (startCol == -1 || endCol == -1) return;

            for (int i = startRow; i < endRow; i++) {

                String testValue = getMergedCellValue(sheet, i, startCol);
                String analyteValue = getMergedCellValue(sheet, i, endCol);
                if(testValue.equals(analyteValue)){
                    mergeCells(sheet, i, startCol, i, endCol);
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    private void mergeRowTestAndAnalyte2(Workbook wb, int sheetIndex, int startRow,int endRow, int maxCol, String testType) {
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            Point testTypeColPoint = selectPosition(sheet, testType, endRow, maxCol);
            Point startColPoint = selectPosition(sheet, "检验检测项目", endRow, maxCol);
            Point endColPoint = selectPosition(sheet, "分析项", endRow, maxCol);

            int testTypeCol = testTypeColPoint.y;
            int startCol = startColPoint.y;
            int endCol = endColPoint.y;

            if (startCol == -1 || endCol == -1) return;

            for (int i = startRow; i < endRow; i++) {
                String testTypeValue = getMergedCellValue(sheet, i, testTypeCol);

                String testValue = getMergedCellValue(sheet, i, startCol);

                String analyteValue = getMergedCellValue(sheet, i, endCol);

                if (StringUtil.isNullOrEmpty(testTypeValue) && !testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, testTypeCol, i, startCol);
                } else if (StringUtil.isNullOrEmpty(testTypeValue) && testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, testTypeCol, i, endCol);
                } else if (!StringUtil.isNullOrEmpty(testTypeValue) && !testTypeValue.equals(testValue) && testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, startCol, i, endCol);
                } else if (!StringUtil.isNullOrEmpty(testTypeValue) && testTypeValue.equals(testValue) && testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, testTypeCol, i, endCol);
                } else if (!StringUtil.isNullOrEmpty(testTypeValue) && testTypeValue.equals(testValue) && !testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, testTypeCol, i, startCol);
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    private  void mergeRowTestAndAnalyte22(Workbook wb, int sheetIndex, int startRow,int endRow, int maxCol, String testType) {
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            Point testTypeColPoint = selectPosition(sheet, testType, endRow, maxCol);
            Point startColPoint = selectPosition(sheet, "检验检测项目", endRow, maxCol);
            Point endColPoint = selectPosition(sheet, "分析项", endRow, maxCol);

            int testTypeCol = testTypeColPoint.y;
            int startCol = startColPoint.y;
            int endCol = endColPoint.y;

            if (startCol == -1 || endCol == -1) return;

            for (int i = startRow; i < endRow; i++) {
                String testTypeValue = getMergedCellValue(sheet, i, testTypeCol);
                String testValue = getMergedCellValue(sheet, i, startCol);
                String analyteValue = getMergedCellValue(sheet, i, endCol);
                if (StringUtil.isNullOrEmpty(testTypeValue) && !testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, endCol, i, testTypeCol);
                } else if (StringUtil.isNullOrEmpty(testTypeValue) && testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, startCol, i, testTypeCol);
                } else if (!StringUtil.isNullOrEmpty(testTypeValue) && !testTypeValue.equals(analyteValue) && testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, startCol, i, endCol);
                } else if (!StringUtil.isNullOrEmpty(testTypeValue) && testTypeValue.equals(analyteValue) && testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, startCol, i, testTypeCol);
                } else if (!StringUtil.isNullOrEmpty(testTypeValue) && testTypeValue.equals(analyteValue) && !testValue.equals(analyteValue)) {
                    mergeCells(sheet, i, endCol, i, testTypeCol);

                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    private String getMergedCellValue(Sheet sheet, int row, int col) {
        try {
            Point p = getSheetMaxRowCol(sheet);
            int endRow = p.x;
            int endCol = p.y;

            if (row > endRow || col > endCol) {
                // Logging can be added here if required
                return "";
            }

            Row currentRow = sheet.getRow(row);
            if (currentRow == null) {
                return "";
            }

            Cell cell = currentRow.getCell(col);
            if (cell == null) {
                return "";
            }

            // Check if the cell is part of a merged region
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress range = sheet.getMergedRegion(i);
                if (range.isInRange(cell)) {
                    // If it's a merged cell, retrieve the value from the top-left cell
                    Row topRow = sheet.getRow(range.getFirstRow());
                    Cell topCell = topRow.getCell(range.getFirstColumn());
                    return getCellValue(topCell);
                }
            }

            // If it's not a merged cell, just retrieve the value
            return getCellValue(cell);
        } catch (Exception e) {
            e.printStackTrace();
            // Logging can be added here if required
        }
        return "";
    }
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            } else {
                return Double.toString(cell.getNumericCellValue());
            }
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            return Boolean.toString(cell.getBooleanCellValue());
        } else if (cell.getCellType() == CellType.FORMULA) {
            // Note: Depending on the formula, you might need to evaluate it first
            return cell.getCellFormula();
        }
        // Handle other cell types if needed
        return "";
    }
    // Utility method to merge cells within a sheet
    private void mergeCells(Sheet sheet, int firstRow, int firstCol, int lastRow, int lastCol) {
        // POI uses 0-based indices for rows and columns
//        firstRow--;
//        firstCol--;
//        lastRow--;
//        lastCol--;

        // Create a CellRangeAddress to represent the region to be merged
        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

        // Add the CellRangeAddress to the sheet
        sheet.addMergedRegion(cellRangeAddress);
    }
    private void mergeRows(Workbook wb, int sheetIndex, int startRow,int endRow, Point unpivotRange, int maxColIndex, String[] unpivotMerge) {
        // 实现合并转置列的逻辑
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            for (int i = startRow; i <= endRow; i++) {
                // Check if we should merge this row based on the unpivotMerge array
                if ((unpivotMerge.length > i - startRow) && ("1".equals(unpivotMerge[i - startRow]))) {
                    // Merge the cells for this row
                    int firstCol = unpivotRange.x - 1; // Convert 1-based index to 0-based
                    int lastCol = unpivotRange.y - 1; // Convert 1-based index to 0-based
                    sheet.addMergedRegion(new CellRangeAddress(i, i, firstCol, lastCol));
                }
            }
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
        }
    }
    private void mergeCells1(Workbook wb, int sheetIndex, int[] colList, int startRow, int endRow) {
        // 实现合并相同检测项目的同列数据的逻辑
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            for (int colIndex : colList) {
                Row tempRowRange = sheet.getRow(startRow);
                if(tempRowRange == null) continue;
                Cell tempCell = tempRowRange.getCell(colIndex);
                if (tempCell == null) continue;
                String tempCellValue = getMergedCellValue(sheet, startRow, colIndex);
                int tempRow = startRow;
                int beforeRow = startRow;

                for (int j = startRow + 1; j <= endRow; j++) {
                    String nowCellValue = getMergedCellValue(sheet, j, colIndex);
                    int[] rangeColumn = getRangeArea(sheet, tempRow, colIndex);
                    int[] rangeColumn2 = getRangeArea(sheet, j, colIndex);

                    boolean shouldMerge = false;
                    if (colIndex == 3) { // For detection project
                        String nowCellValue22 = getMergedCellValue(sheet, j, 2);
                        String nowCellValue33 = getMergedCellValue(sheet, tempRow, 2);
                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1] && nowCellValue22.equals(nowCellValue33);
                    } else {
                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1];
                    }

                    if (shouldMerge) {
                        beforeRow++;
                    } else {
                        if (tempRow < beforeRow) {
                            CellRangeAddress range = new CellRangeAddress(tempRow, beforeRow, colIndex, colIndex);
                            sheet.addMergedRegion(range);
                        }
                        tempCellValue = nowCellValue;
                        tempRow = j;
                        beforeRow++;
                    }
                }

//                // Merge the last set of cells if needed
//                if (tempRow < beforeRow) {
//                    CellRangeAddress range = new CellRangeAddress(tempRow, beforeRow, colIndex, colIndex);
//                    sheet.addMergedRegion(range);
//                }
            }
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
        }
    }
    private void mergeCells3(Workbook wb, int sheetIndex, int[] colList, int startRow, int endRow) {
        // 实现合并相同检测项目的同列数据的逻辑
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            for (int colIndex : colList) {
                Row tempRowRange = sheet.getRow(startRow);
                if(tempRowRange == null) continue;
                Cell tempCell = tempRowRange.getCell(colIndex);
                if (tempCell == null) continue;
                String tempCellValue = getMergedCellValue(sheet, startRow, colIndex);

                int tempRow = startRow;
                int beforeRow = startRow;

                for (int j = startRow + 1; j <= endRow; j++) {
                    String nowCellValue = getMergedCellValue(sheet, j, colIndex);
                    int[] rangeColumn = getRangeArea(sheet, tempRow, colIndex);
                    int[] rangeColumn2 = getRangeArea(sheet, j, colIndex);


                    boolean shouldMerge = false;
                    if (colIndex == 3) { //表示检测项目
                        String nowCellValue22 = getMergedCellValue(sheet, j, 1);

                        String nowCellValue33 = getMergedCellValue(sheet, tempRow, 1);
                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1] && nowCellValue22.equals(nowCellValue33);
                    } else if (colIndex == 4) { //表示分析项目
                        String nowCellValue22 = getMergedCellValue(sheet, j, 3);

                        String nowCellValue33 = getMergedCellValue(sheet, tempRow, 3);

                        String nowCellValue222 = getMergedCellValue(sheet, j, 1);

                        String nowCellValue333 = getMergedCellValue(sheet, tempRow, 1);

                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1] && nowCellValue22.equals(nowCellValue33) && nowCellValue222.equals(nowCellValue333);
                    } else {
                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1];
                    }

                    if (shouldMerge) {
                        beforeRow++;
                    } else {
                        if (tempRow < beforeRow) {
                            CellRangeAddress range = new CellRangeAddress(tempRow, beforeRow, colIndex, colIndex);
                            sheet.addMergedRegion(range);
                        }
                        tempCellValue = nowCellValue;
                        tempRow = j;
                        beforeRow++;
                    }
                }

//                // Merge the last set of cells if needed
//                if (tempRow < beforeRow) {
//                    CellRangeAddress range = new CellRangeAddress(tempRow, beforeRow, colIndex, colIndex);
//                    sheet.addMergedRegion(range);
//                }
            }
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
        }
    }
    private  void mergeCells33(Workbook wb, int sheetIndex, int[] colList, int startRow, int endRow) {
        // 实现合并相同检测项目的同列数据的逻辑
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            for (int colIndex : colList) {
                Row tempRowRange = sheet.getRow(startRow);
                if(tempRowRange == null) continue;
                Cell tempCell = tempRowRange.getCell(colIndex);
                if (tempCell == null) continue;
                String tempCellValue = getMergedCellValue(sheet, startRow, colIndex);
                int tempRow = startRow;
                int beforeRow = startRow;

                for (int j = startRow + 1; j <= endRow; j++) {
                    String nowCellValue = getMergedCellValue(sheet, j, colIndex);
                    int[] rangeColumn = getRangeArea(sheet, tempRow, colIndex);
                    int[] rangeColumn2 = getRangeArea(sheet, j, colIndex);

                    boolean shouldMerge = false;
                    if (colIndex == 3) { // For detection project
                        //检测项目值（写死了）
                        String nowCellValue22 = getMergedCellValue(sheet, j, 2);
                        //检测项目值（写死了）
                        String nowCellValue33 = getMergedCellValue(sheet, tempRow, 2);

                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1] && nowCellValue22.equals(nowCellValue33);
                    } else if (colIndex == 4) { // For analysis project
                        String nowCellValue22 = getMergedCellValue(sheet, j, 3);
                        String nowCellValue33 = getMergedCellValue(sheet, tempRow, 3);
                        String nowCellValue222 = getMergedCellValue(sheet, j, 2);
                        String nowCellValue333 = getMergedCellValue(sheet, tempRow, 2);
                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1] && nowCellValue22.equals(nowCellValue33) && nowCellValue222.equals(nowCellValue333);
                    } else {
                        shouldMerge = tempCellValue.equals(nowCellValue) && rangeColumn2[1] == rangeColumn[1];
                    }

                    if (shouldMerge) {
                        beforeRow++;
                    } else {
                        if (tempRow < beforeRow) {
                            CellRangeAddress range = new CellRangeAddress(tempRow, beforeRow, colIndex, colIndex);
                            sheet.addMergedRegion(range);
                        }
                        tempCellValue = nowCellValue;
                        tempRow = j;
                        beforeRow++;
                    }
                }

//                // Merge the last set of cells if needed
//                if (tempRow < beforeRow) {
//                    CellRangeAddress range = new CellRangeAddress(tempRow, beforeRow, colIndex, colIndex);
//                    sheet.addMergedRegion(range);
//                }
            }
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
        }
    }
    private int[] getRangeArea(Sheet sheet, int row, int col) {
        try{
            Point p = getSheetMaxRowCol(sheet);
            int endRow = p.x;
            int endCol = p.y;
            if (row > endRow || col > endCol)
            {
                return null;
            }
            int[] rangeArea = new int[]{1, 1}; // Default to 1x1 if the cell is not merged
            int rowIndex = row ; // Convert to 0-based index
            int colIndex = col ; // Convert to 0-based index

            // Loop through all merged regions to find one that contains the specified cell
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                if (mergedRegion.isInRange(rowIndex, colIndex)) {
                    rangeArea[0] = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                    rangeArea[1] = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
                    break;
                }
            }

            return rangeArea;
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
        }
        return null;
    }
    private Workbook setAutoRowHeight(Workbook wb, int sheetIndex, int startRow, int startCol, int maxColIndex) {
        // 实现设置行高的逻辑
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            Row row = sheet.getRow(startRow);
            if (row != null) {
                for (int colIndex = startCol; colIndex <= maxColIndex; colIndex++) {
                    sheet.autoSizeColumn(colIndex,true);
                }
            }
        } catch (Exception ex) {
            // Handle exceptions here
            ex.printStackTrace();
        }
        return wb;
    }
    private void dealMergedAreaInPages_new(Workbook wb, int sheetIndex,int startRow ,int endCol, int[] colseq) {
        try {
            Sheet sheet = wb.getSheetAt(sheetIndex);
            int endRow = getSheetMaxRowCol(sheet).x;

            Point p = selectPosition(sheet, "检验检测项目", endRow, endCol);
            int testNoIndex = p.y;
            if (testNoIndex == -1) return;

            int maxColIndex = endCol;

            List<Integer> firstPageRowIndex = getNewPageFirstRow(sheet);

            if (firstPageRowIndex.isEmpty()) {
                int[] iRange = getRepeatingRowsRange(sheet, endCol);

                if (iRange == null || startRow == iRange[1]) {
                    lastPageFirstRow = 1;
                } else {
                    lastPageFirstRow = iRange[1] + 1;
                }
            } else {
                lastPageFirstRow = firstPageRowIndex.get(firstPageRowIndex.size() - 1);
            }

            for (int i : firstPageRowIndex) {
                System.out.println("i = " + i);
                if (i <= endRow && i > p.x) {
                    dealMergedBetweenPages(sheet, i, maxColIndex, colseq);
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    /// 获取sheet每一页的第一行索引,第一页除外
    private List<Integer> getNewPageFirstRow(Sheet sheet) {
        List<Integer> arrayFr = new ArrayList<>();
        try {
            // 获取默认行分页符集合
            int[] pageBreaks = sheet.getRowBreaks();

            // 遍历分页符，获取每个分页符位置的行号
            for (int i = 1; i < pageBreaks.length; i++) { // 从第二个分页符开始
                arrayFr.add(pageBreaks[i]);
            }
            return arrayFr;
        } catch (Exception ex) {
            ex.printStackTrace();
            return arrayFr;
        }
    }
    /// <summary>
    /// 返回sheet重复区域,分别为起始行号,结束行号,起始列号,结束列号
    /// </summary>
    /// <param name="sheet">工作表对象</param>
    /// <param name="endCol">结束列号</param>
    /// <returns>整型数组,分别为起始行号,结束行号,起始列号,结束列号</returns>
    private int[] getRepeatingRowsRange(Sheet sheet, int endCol) {
        String printTitleRows = sheet.getRepeatingRows() != null ? sheet.getRepeatingRows().formatAsString() : null;

        String printTitleColumns = sheet.getRepeatingColumns() != null ? sheet.getRepeatingColumns().formatAsString() : null;

        if (printTitleRows == null || printTitleRows.isEmpty()) {
            return null;
        }

        int[] ir = new int[]{-1, -1, -1, -1};

        // Parse printTitleRows
        String[] rowRange = printTitleRows.split(":");
        ir[0] = Integer.parseInt(rowRange[0].replaceAll("\\D", ""));
        ir[1] = Integer.parseInt(rowRange[1].replaceAll("\\D", ""));

        // Parse printTitleColumns if available
        if (printTitleColumns == null || printTitleColumns.isEmpty()) {
            ir[2] = 1;
            ir[3] = endCol;
        } else {
            String[] colRange = printTitleColumns.split(":");
            ir[2] = convertColumnLetterToNumber(colRange[0].replaceAll("\\d", ""));
            ir[3] = convertColumnLetterToNumber(colRange[1].replaceAll("\\d", ""));
        }

        return ir;
    }
    // Convert Excel column letter to number
    private int convertColumnLetterToNumber(String letter) {
        int column = 0;
        for (int i = 0; i < letter.length(); i++) {
            column = column * 26 + (letter.charAt(i) - 'A' + 1);
        }
        return column;
    }
    /// <summary>
    /// 处理2页之间的合并单元格,上下单独合并,不需要合并列的列不用处理
    /// </summary>
    /// <param name="sheet">工作表对象</param>
    /// <param name="rowIndex">页首行行号</param>
    /// <param name="endCol">最大行</param>
    /// <param name="colseq">需要合并的列</param>
    private void dealMergedBetweenPages(Sheet sheet, int rowIndex, int endCol, int[] colseq) {
        try {
            int i = 0;
            while (i < colseq.length && i <= endCol) {
                Cell cell = sheet.getRow(rowIndex).getCell(colseq[i]);
                if (cell == null) continue;
                boolean isMerged = isCellMerged(sheet, rowIndex, colseq[i]);
                if (isMerged) {
                    // Get the merged region without splitting. If it needs to span pages, then split.
                    int[] mergedArea = getMergedArea(sheet, rowIndex, colseq[i], false);

                    // If the start row of the merged area is greater than the page start row, it does not need to be processed.
                    if (mergedArea[0] >= rowIndex) {
                        i++;
                        continue;
                    } else {
                        // Get the value of the merged area. After splitting and merging, it needs to be reassigned.
                        String cellValue = getMergedCellValue(sheet, rowIndex, colseq[i]);

                        // This represents the need to split across pages.
                        unmergeCell(sheet, mergedArea, cellValue);

                        // Merge the previous page.
                        mergeCells(sheet, mergedArea[0], mergedArea[1], rowIndex - 1, mergedArea[3]);
                        //todo:设置边框为全框线

                        // Merge the next page.
                        mergeCells(sheet, rowIndex, mergedArea[1], mergedArea[2], mergedArea[3]);
                        //todo:设置边框为全框线

                        // Should jump outside the merged area,
                        // Check the end column of the merged area and find the first merged column array index that is larger than it.
                        boolean jump = false;
                        for (int j = i + 1; j < colseq.length; j++) {
                            if (colseq[j] > mergedArea[3]) {
                                i = j;
                                jump = true;
                                break;
                            }
                        }
                        if (jump) {
                            continue;
                        } else {
                            i++;
                            continue;
                        }
                    }
                } else {
                    i++;
                    continue;
                }
            }
        } catch (Exception e) {
            // Log the exception as per your logging mechanism
            e.printStackTrace();
        }
    }
    private boolean isCellMerged(Sheet sheet, int rowIndex, int columnIndex) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                return true;
            }
        }
        return false;
    }
    private int[] getMergedArea(Sheet sheet, int nRow, int nCol, boolean isUnMerge) {
        int[] mergedArea = {0, 0, 0, 0};
        try {
            Cell cell = sheet.getRow(nRow).getCell(nCol);
            int sheetMergeCount = sheet.getNumMergedRegions();

            for (int i = 0; i < sheetMergeCount; i++) {
                CellRangeAddress range = sheet.getMergedRegion(i);
                if (range.isInRange(nRow, nCol)) {
                    mergedArea[0] = range.getFirstRow();
                    mergedArea[1] = range.getFirstColumn();
                    mergedArea[2] = range.getLastRow();
                    mergedArea[3] = range.getLastColumn();

                    if (isUnMerge) {
                        String cellValue = getMergedCellValue(sheet, range.getFirstRow(), range.getFirstColumn());
                        sheet.removeMergedRegion(i);
                        // After unmerge, set the value to the first cell of the previously merged area.
                        Row row = sheet.getRow(mergedArea[0]);
                        if (row == null) {
                            row = sheet.createRow(mergedArea[0]);
                        }
                        Cell firstCell = row.getCell(mergedArea[1]);
                        if (firstCell == null) {
                            firstCell = row.createCell(mergedArea[1]);
                        }
                        firstCell.setCellValue(cellValue);
                    }
                    break;
                }
            }
        } catch (Exception ex) {
            // Log the exception as per your logging mechanism
            ex.printStackTrace();
        }
        return mergedArea;
    }
    private void unmergeCell(Sheet sheet, int[] mergedArea, String cellValue) {
        int firstRow = mergedArea[0];
        int firstCol = mergedArea[1];
        int lastRow = mergedArea[2];
        int lastCol = mergedArea[3];

        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() == firstRow && mergedRegion.getFirstColumn() == firstCol &&
                    mergedRegion.getLastRow() == lastRow && mergedRegion.getLastColumn() == lastCol) {
                sheet.removeMergedRegion(i);
                // After unmerging, set the value to all cells that were part of the merged region.
                for (int rowIdx = firstRow; rowIdx <= lastRow; rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) {
                        row = sheet.createRow(rowIdx);
                    }
                    for (int colIdx = firstCol; colIdx <= lastCol; colIdx++) {
                        Cell cell = row.getCell(colIdx);
                        if (cell == null) {
                            cell = row.createCell(colIdx);
                        }
                        cell.setCellValue(cellValue);
                    }
                }
                break; // Assuming only one merged region matches the criteria.
            }
        }
    }
    private void mergeTestCell(Workbook wb, int sheetIndex, String str1, String str2, int endCol) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        int endRow = getSheetMaxRowCol(sheet).x;

        Point p1 = selectPosition(sheet, str1, endRow, endCol);
        Point p2 = selectPosition(sheet, str2, endRow, endCol);

        if (p1.x == -1 || p1.y == -1 || p2.x == -1 || p2.y == -1 ) return;
        // 不在同一行不合并
        if (p1.x != p2.x) return;
        // 不是相邻单元格不合并
        if (p1.y + 1 != p2.y) return;
        else {
            CellRangeAddress rangeProgram = new CellRangeAddress(p1.x , p2.x , p1.y , p2.y );
            sheet.addMergedRegion(rangeProgram);

        }
    }
    private void mergeTestCell2(Workbook wb, int sheetIndex, String str3, String str1, String str2, int endCol) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        int endRow = getSheetMaxRowCol(sheet).x;

        Point p3 = selectPosition(sheet, str3, endRow, endCol);
        Point p1 = selectPosition(sheet, str1, endRow, endCol);
        Point p2 = selectPosition(sheet, str2, endRow, endCol);


        if (p1.x == -1 || p1.y == -1 || p2.x == -1 || p2.y == -1 || p3.x == -1 || p3.y == -1) return;
        // 不在同一行不合并
        if (p1.x != p2.x) return;
        // 不是相邻单元格不合并
        if (p1.y + 1 != p2.y) return;
        else {
            // 原来的值
            String testNoValue = getMergedCellValue(sheet, p1.x, p1.y);

            CellRangeAddress rangeProgram = new CellRangeAddress(p3.x , p2.x , p3.y , p2.y);
            sheet.addMergedRegion(rangeProgram);

            // 改回原来的值，因为测试分类列在检验检测项目前面
            Row row = sheet.getRow(p3.x );
            Cell cell = row.getCell(p3.y );
            if (cell == null) {
                cell = row.createCell(p3.y);
            }
            cell.setCellValue(testNoValue);
        }
    }

    private void mergeTestCell22(Workbook wb, int sheetIndex, String str3, String str1, String str2, int endCol) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        int endRow = getSheetMaxRowCol(sheet).x;

        Point p3 = selectPosition(sheet, str3, endRow, endCol);
        Point p1 = selectPosition(sheet, str1, endRow, endCol);
        Point p2 = selectPosition(sheet, str2, endRow, endCol);

        if (p1.x == -1 || p1.y == -1 || p2.x == -1 || p2.y == -1 || p3.x == -1 || p3.y == -1) return;
        // 不在同一行不合并
        if (p1.x != p2.x) return;
        // 不是相邻单元格不合并
        if (p1.y + 1 != p2.y) return;
        else {
            // 原来的值
            String testNoValue = getMergedCellValue(sheet, p1.x, p1.y);

            CellRangeAddress rangeProgram = new CellRangeAddress(p1.x , p3.x , p1.y , p3.y);
            sheet.addMergedRegion(rangeProgram);

            // 改回原来的值，因为测试分类列在检验检测项目前面
            Row row = sheet.getRow(p3.x);
            Cell cell = row.getCell(p3.y );
            if (cell == null) {
                cell = row.createCell(p3.y);
            }
            cell.setCellValue(testNoValue);
        }
    }
    private String saveWorkBookAsPdf(Workbook wb, String saveDir) {
       return null;
    }
    private int getNumberOfColumns(Sheet sheet) {
        int maxColumns = 0;
        for (Row row : sheet) {
            if (row.getPhysicalNumberOfCells() > maxColumns) {
                maxColumns = row.getPhysicalNumberOfCells();
            }
        }
        return maxColumns;
    }
//    /**
//     * 在指定行之前插入一行并垂直添加图片
//     *
//     * @param workbookPath Excel文件路径
//     * @param sheetIndex   工作表索引
//     * @param imagePaths   图片文件路径数组
//     * @param rowIndex     在这一行之前插入新行并添加图片
//     * @return true if successful, false otherwise
//     */
//    public boolean addImagesToRowWithAdjustedHeight(String workbookPath, int sheetIndex, String[] imagePaths, int rowIndex) {
//        boolean success = true;
//        FileInputStream fileInputStream = null;
//        FileOutputStream fileOutputStream = null;
//        Workbook workbook = null;
//
//        try {
//            fileInputStream = new FileInputStream(workbookPath);
//            workbook = WorkbookFactory.create(fileInputStream);
//            fileInputStream.close(); // Close the FileInputStream immediately after use
//
//            Sheet sheet = workbook.getSheetAt(sheetIndex);
//            Point pSheet = getSheetMaxRowCol(sheet);
//            int maxRowIndex = pSheet.x -1 ;
//            int maxColIndex = pSheet.y -1 ;
//            System.out.println("maxColIndex = " + maxColIndex);
//            Drawing<?> drawing = sheet.createDrawingPatriarch(); // Create a single Drawing patriarch for the whole operation
//            // 创建边框样式
//            CellStyle borderStyle = workbook.createCellStyle();
//            borderStyle.setBorderTop(BorderStyle.THIN);
//            borderStyle.setBorderBottom(BorderStyle.THIN);
//            borderStyle.setBorderLeft(BorderStyle.THIN);
//            borderStyle.setBorderRight(BorderStyle.THIN);
//            for (int i = 0; i < imagePaths.length; i++) {
//                String imagePath = imagePaths[i];
//                System.out.println("imagePath = " + imagePath);
//
//                // Shift rows down to make space for the new row with the image
//                sheet.shiftRows(rowIndex, sheet.getLastRowNum(), 1);
//
//                // Create a new row at the specified index
//                Row row = sheet.createRow(rowIndex);
//                // 合并单元格，跨越5列
//                sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, maxColIndex));
//                // 设置合并后的单元格的边框
//                for (int col = 0; col <= maxColIndex; col++) {
//                    Cell cell = row.createCell(col);
//                    cell.setCellStyle(borderStyle);
//                }
//                // Read the image file into a byte array
//                InputStream inputStream = new FileInputStream(imagePaths[i]);
//                byte[] bytes = IOUtils.toByteArray(inputStream);
//                inputStream.close();
//
//                // Add the picture to the workbook
//                int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
//                CreationHelper helper = workbook.getCreationHelper();
//                ClientAnchor anchor = helper.createClientAnchor();
//
//                // Set the anchor points for the image
//                anchor.setCol1(0); // 图片从第一列开始
//                anchor.setRow1(rowIndex); // 图片开始的行
//                anchor.setCol2(maxColIndex); // 图片结束的列，如果图片需要跨多列则需要调整
//                anchor.setRow2(rowIndex + 1); // 图片结束的行，这里设置为下一行，因为图片只占据一行
//                // anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
//
//                // Create the picture and manually set the size
//                Picture pict = drawing.createPicture(anchor, pictureIdx);
//
//                // pict.resize();
//// Get the image dimensions
//                double imageHeight = pict.getImageDimension().getHeight();
//                double imageWidth = pict.getImageDimension().getWidth();
//
//                // Convert the image dimensions to points
//                float rowHeightInPoints = (float) Units.pixelToPoints(imageHeight);
//                double colWidthInPoints = Units.pixelToPoints(imageWidth) / 5; // Adjust width to fit 5 columns
//
//                // Adjust the row height to fit the image
//                row.setHeightInPoints(rowHeightInPoints + 1);
//                // Increase rowIndex for the next image
//                rowIndex++;
//            }
//
//            // Write changes back to the workbook
//            fileOutputStream = new FileOutputStream(workbookPath);
//            workbook.write(fileOutputStream);
//
//        } catch (Exception e) {
//            e.printStackTrace();
//            success = false;
//        } finally {
//            try {
//                if (fileOutputStream != null) fileOutputStream.close();
//                if (workbook != null) workbook.close();
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//        }
//        return success;
//    }

    // 生成条形码
    public void generateBarcode(String text, String filePath, int width, int height) {
        Path path = Paths.get(filePath);
        // 检查文件是否存在
        if (Files.exists(path)) {
            return;
        }
        BitMatrix bitMatrix = null;
        try {
            bitMatrix = new MultiFormatWriter().encode(text, BarcodeFormat.CODE_128, width, height);
        } catch (WriterException e) {
            throw new RuntimeException(e);
        }

        try {
            MatrixToImageWriter.writeToPath(bitMatrix, "PNG", path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void generateQRCode(String text, String filePath, int width, int height) {
        Path path = Paths.get(filePath);
        // 检查文件是否存在
        if (Files.exists(path)) {
            return;
        }

        BitMatrix bitMatrix = null;
        try {
            bitMatrix = new MultiFormatWriter().encode(text, BarcodeFormat.QR_CODE, width, height);
        } catch (WriterException e) {
            throw new RuntimeException(e);
        }
        try {
            MatrixToImageWriter.writeToPath(bitMatrix, "PNG", path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    //指定图片在excel的始末单元格，按照dx，dy进行移动
//    public void insertImageIntoExcel(String excelFilePath, int sheetIndex, String[] imagesPath,int[][] imagePoints) {
//        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFilePath))) {
//            // 打开Excel文件
//            Sheet sheet = workbook.getSheetAt(sheetIndex);
//            // 获取画图管理器
//            Drawing<?> drawing = sheet.createDrawingPatriarch();
//            CreationHelper helper = workbook.getCreationHelper();
//
//            for (int i = 0; i < imagesPath.length; i++) {
//                String imagePath = imagesPath[i];
//                if (new File(imagePath).exists()){
//                    InputStream inputStream = new FileInputStream(imagePath);
//                    byte[] bytes = IOUtils.toByteArray(inputStream);
//                    inputStream.close();
//
//                    ClientAnchor anchor = helper.createClientAnchor();
//                    // 设置图片位置和大小
//                    anchor.setCol1(imagePoints[i][0]);
//                    anchor.setRow1(imagePoints[i][1]);
//                    anchor.setCol2(imagePoints[i][2]);
//                    anchor.setRow2(imagePoints[i][3]);
//                    anchor.setDx1(imagePoints[i][4]);
//                    anchor.setDy1(imagePoints[i][5]);
//                    anchor.setDx2(imagePoints[i][6]);
//                    anchor.setDy2(imagePoints[i][7]);
//                    anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
//
//                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
//                    drawing.createPicture(anchor, pictureIdx);
//                }else{
//                    System.out.println("Image not found: " + imagePath);
//                }
//            }
//            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
//                workbook.write(fileOutputStream);
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//    //从某个单元格开始，按图片实际大小，等比例缩放
//    public  void insertImageIntoExcel(String excelFilePath, int sheetIndex, String[] imagesPath,int[][] imagePoints,double scaleFactor) {
//        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFilePath))) {
//
//            Sheet sheet = workbook.getSheetAt(sheetIndex);
//            // 获取画图管理器
//            Drawing<?> drawing = sheet.createDrawingPatriarch();
//            CreationHelper helper = workbook.getCreationHelper();
//
//            for (int i = 0; i < imagesPath.length; i++) {
//                String imagePath = imagesPath[i];
//                if (new File(imagePath).exists()){
//                    InputStream inputStream = new FileInputStream(imagePath);
//                    byte[] bytes = IOUtils.toByteArray(inputStream);
//                    inputStream.close();
//                    BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imagePath));
//                    //图片的真实长宽
//                    int imageWidth = bufferedImage.getWidth();
//                    int imageHeight = bufferedImage.getHeight();
//                    // 获取图片起始单元格
//                    int col1 = imagePoints[i][0];
//                    int row1 = imagePoints[i][1];
//                    // 计算单元格的长宽
//                    float cellWidth = sheet.getColumnWidthInPixels(col1);
//                    float cellHeight = sheet.getRow(row1).getHeightInPoints() / 72 * 96; // 高度转换为像素
//                    // 计算需要的长宽比例的系数
//                    double a = imageWidth * scaleFactor / cellWidth;
//                    double b = imageHeight * scaleFactor / cellHeight;
//                    ClientAnchor anchor = helper.createClientAnchor();
//
//                    anchor.setCol1(col1);
//                    anchor.setRow1(row1);
//
//                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
//                    Picture pict = drawing.createPicture(anchor, pictureIdx);
//                    pict.resize(a,b);
//                }else{
//                    System.out.println("Image not found: " + imagePath);
//                }
//
//            }
//
//            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
//                workbook.write(fileOutputStream);
//            }
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//    public  void insertImageIntoExcel2(String excelFilePath, int sheetIndex, String[] imagesPath,int[][] imagePoints,double scaleFactor) {
//        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFilePath))) {
//
//            Sheet sheet = workbook.getSheetAt(sheetIndex);
//            // 获取画图管理器
//            Drawing<?> drawing = sheet.createDrawingPatriarch();
//            CreationHelper helper = workbook.getCreationHelper();
//
//            for (int i = 0; i < imagesPath.length; i++) {
//                String imagePath = imagesPath[i];
//                if (new File(imagePath).exists()){
//                    InputStream inputStream = new FileInputStream(imagePath);
//                    byte[] bytes = IOUtils.toByteArray(inputStream);
//                    inputStream.close();
//
//                    ClientAnchor anchor = helper.createClientAnchor();
//                    // 设置图片位置和大小
//                    anchor.setCol1(imagePoints[i][0]);
//                    anchor.setRow1(imagePoints[i][1]);
//                    anchor.setCol2(imagePoints[i][2]);
//                    anchor.setRow2(imagePoints[i][3]);
////                    anchor.setDx1(imagePoints[i][4]);
////                    anchor.setDy1(imagePoints[i][5]);
////                    anchor.setDx2(imagePoints[i][6]);
////                    anchor.setDy2(imagePoints[i][7]);
////                    anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
//
//                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
//                    Picture pict = drawing.createPicture(anchor, pictureIdx);
//
//                    BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imagePath));
//                    //图片的真实长宽
//                    int imageWidth = bufferedImage.getWidth();
//                    int imageHeight = bufferedImage.getHeight();
//                    System.out.println("imageWidth = " + imageWidth);
//                    System.out.println("imageHeight = " + imageHeight);
//                    // 目标大小（像素）
//                    // 目标大小（像素），根据你设定的范围计算
//                    double targetWidth = (anchor.getCol2() - anchor.getCol1()) * sheet.getColumnWidthInPixels(1);
//                    double targetHeight = (anchor.getRow2() - anchor.getRow1()) * sheet.getDefaultRowHeightInPoints() / 72 * 96;
//
//                    // 计算比例
//                    double widthRatio = targetWidth / imageWidth;
//                    double heightRatio = targetHeight / imageHeight;
//
//                    // 使用最小的比例来保持图片的纵横比
//                    double resizeRatio = Math.min(widthRatio, heightRatio);
//                    System.out.println("resizeRatio = " + resizeRatio);
//                    pict.resize();
//                    //pict.resize(resizeRatio);
//                   }else{
//                    System.out.println("Image not found: " + imagePath);
//                }
//
//            }
//
//            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
//                workbook.write(fileOutputStream);
//            }
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

    public void insertImages(String localsyPath, int sheetIndex,String[] imagesPath,int[][] imagePoints,double[] scaleFactors) {
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(localsyPath))) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            CreationHelper helper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            for (int i = 0; i < imagesPath.length; i++) {
                String imagePath = imagesPath[i];
                if (new File(imagePath).exists()) {
                    InputStream inputStream = new FileInputStream(imagePath);
                    byte[] bytes = IOUtils.toByteArray(inputStream);
                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                    inputStream.close();
                    BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imagePath));
                    //图片的真实长宽
                    int pictureWidth = bufferedImage.getWidth();
                    int pictureHeight = bufferedImage.getHeight();
                    // 获取图片起始单元格
                    int col1 = imagePoints[i][0];
                    int row1 = imagePoints[i][1];
                    int col2 = imagePoints[i][2];
                    int row2 = imagePoints[i][3];
                    // 计算单元格的总宽度
                    float cellWidth = 0;
                    for (int col = col1; col <= col2; col++) {
                        cellWidth += sheet.getColumnWidthInPixels(col);
                    }
                    // 计算单元格的总高度
                    float cellHeight = 0;
                    for (int r = row1; r <= row2; r++) {
                        Row row = sheet.getRow(r);
//                        if (row == null) {
//                            row = sheet.createRow(r); // 创建一个新的空行
//                        }
//                        cellHeight += row.getHeightInPoints() / 72 * 96; // 转换为像素
                        if (row != null) {  // 检查 row 是否为 null
                            cellHeight += row.getHeightInPoints() / 72 * 96; // 转换为像素
                        } else {
                            System.out.println("insertImages:Row " + r + " is null.");
                            // 如果该行为空，使用默认高度
                            float defaultHeight = sheet.getDefaultRowHeightInPoints(); // 获取默认行高
                            cellHeight += defaultHeight / 72 * 96; // 转换为像素
                        }
                    }
                    // 计算缩放比例，保持图片的宽高比一致
                    //double scale = Math.min(cellWidth / pictureWidth, cellHeight / pictureHeight);

                    double scale = 0.32;
                    // 计算缩放后的图片宽高
                    int newWidth = (int) (pictureWidth * scale);
                    int newHeight = (int) (pictureHeight * scale);

                    // 计算偏移量，使图片居中
                    //int dx1 = (int) ((cellWidth - newWidth) / 2);
                    //int dy1 = (int) ((cellHeight - newHeight) / 2);
                    //int dx2 = dx1 + newWidth;
                    //int dy2 = dy1 + newHeight;
                    //// 设置锚点并插入图片
                    //ClientAnchor anchor = helper.createClientAnchor();
                    //anchor.setCol1(col1);
                    //anchor.setRow1(row1);
                    //Picture picture = drawing.createPicture(anchor, pictureIdx);
                    //picture.resize(scale);

                    // 计算偏移量，使图片居中
                    int dx1 = 0;
                    int dy1 = 0;
                    int dx2 = dx1 + newWidth;
                    int dy2 = dy1 + newHeight;
                    // 设置锚点并插入图片
                    ClientAnchor anchor = helper.createClientAnchor();
                    anchor.setCol1(col1);
                    anchor.setRow1(row1);
                    anchor.setDx1(dx1 * Units.EMU_PER_PIXEL);
                    anchor.setDy1(dy1 * Units.EMU_PER_PIXEL);
                    anchor.setDx2(dx2 * Units.EMU_PER_PIXEL);
                    anchor.setDy2(dy2 * Units.EMU_PER_PIXEL);
                    anchor.setCol2(col2);
                    anchor.setRow2(row2);
                    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                    Picture picture = drawing.createPicture(anchor, pictureIdx);
                    //确定了dx1，dy1，dx2，dy2。就不需要resize
                    picture.resize(scale * 1.58,scale * 1.59);
                } else {
                    System.out.println("Image not found: " + imagePath);
                }

            }

            try (FileOutputStream fileOutputStream = new FileOutputStream(localsyPath)) {
                workbook.write(fileOutputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public void insertImagesCenter(String localsyPath, int sheetIndex,String[] imagesPath,int[][] imagePoints,double[] scaleFactors) {
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(localsyPath))) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            CreationHelper helper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            for (int i = 0; i < imagesPath.length; i++) {
                String imagePath = imagesPath[i];
                if (new File(imagePath).exists()) {
                    InputStream inputStream = new FileInputStream(imagePath);
                    byte[] bytes = IOUtils.toByteArray(inputStream);
                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                    inputStream.close();
                    BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imagePath));
                    //图片的真实长宽
                    int pictureWidth = bufferedImage.getWidth();
                    int pictureHeight = bufferedImage.getHeight();
                    // 获取图片起始单元格
                    int col1 = imagePoints[i][0];
                    int row1 = imagePoints[i][1];
                    int col2 = imagePoints[i][2];
                    int row2 = imagePoints[i][3];
                    // 计算单元格的总宽度
                    float cellWidth = 0;
                    for (int col = col1; col <= col2; col++) {
                        cellWidth += sheet.getColumnWidthInPixels(col);
                    }
                    // 计算单元格的总高度
                    float cellHeight = 0;
                    for (int r = row1; r <= row2; r++) {
                        Row row = sheet.getRow(r);
//                        if (row == null) {
//                            row = sheet.createRow(r); // 创建一个新的空行
//                        }
//                        cellHeight += row.getHeightInPoints() / 72 * 96; // 转换为像素
                        if (row != null) {  // 检查 row 是否为 null
                            cellHeight += row.getHeightInPoints() / 72 * 96; // 转换为像素
                        } else {
                            System.out.println("insertImagesCenter:Row " + r + " is null.");
                            // 如果该行为空，使用默认高度
                            float defaultHeight = sheet.getDefaultRowHeightInPoints(); // 获取默认行高
                            cellHeight += defaultHeight / 72 * 96; // 转换为像素
                        }
                    }
                    // 计算缩放比例，保持图片的宽高比一致
                    //double scale = Math.max(cellWidth / pictureWidth, cellHeight / pictureHeight);
                    double scale = Math.min(cellWidth / pictureWidth, cellHeight / pictureHeight);

                    scale = scale * scaleFactors[i];
                    // 计算缩放后的图片宽高
                    int newWidth = (int) (pictureWidth * scale);
                    int newHeight = (int) (pictureHeight * scale);

                    // 计算偏移量，使图片居中
                    int dx1 = (int) ((cellWidth - newWidth) / 2);
                    int dy1 = (int) ((cellHeight - newHeight) / 2);
                    //int dx2 = dx1 + newWidth;
                    //int dy2 = dy1 + newHeight;
                    // 设置锚点并插入图片
                    ClientAnchor anchor = helper.createClientAnchor();
                    anchor.setCol1(col1);
                    anchor.setRow1(row1);
                    anchor.setDx1(dx1 * Units.EMU_PER_PIXEL);
                    anchor.setDy1(dy1 * Units.EMU_PER_PIXEL);
                    //anchor.setDx2(dx2 * Units.EMU_PER_PIXEL);
                    //anchor.setDy2(dy2 * Units.EMU_PER_PIXEL);
                    //anchor.setCol2(mergedRegion.getLastColumn());
                    //anchor.setRow2(mergedRegion.getLastRow());
                    //anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                    Picture picture = drawing.createPicture(anchor, pictureIdx);
                    picture.resize(scale);
                } else {
                    System.out.println("Image not found: " + imagePath);
                }

            }

            try (FileOutputStream fileOutputStream = new FileOutputStream(localsyPath)) {
                workbook.write(fileOutputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public  void insertImageIntoExcel(String excelFilePath, int sheetIndex, String[] imagesPath,int[][] imagePoints) {
        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFilePath))) {
            // 打开Excel文件
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            // 获取画图管理器
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            CreationHelper helper = workbook.getCreationHelper();

            for (int i = 0; i < imagesPath.length; i++) {
                String imagePath = imagesPath[i];
                if (new File(imagePath).exists()){
                    InputStream inputStream = new FileInputStream(imagePath);
                    byte[] bytes = IOUtils.toByteArray(inputStream);

                    inputStream.close();

                    ClientAnchor anchor = helper.createClientAnchor();
                    // 设置图片位置和大小
                    anchor.setCol1(imagePoints[i][0]);
                    anchor.setRow1(imagePoints[i][1]);
                    anchor.setCol2(imagePoints[i][2]);
                    anchor.setRow2(imagePoints[i][3]);
                    anchor.setDx1(imagePoints[i][4] * Units.EMU_PER_PIXEL);
                    anchor.setDy1(imagePoints[i][5] * Units.EMU_PER_PIXEL);
                    anchor.setDx2(imagePoints[i][6] * Units.EMU_PER_PIXEL);
                    anchor.setDy2(imagePoints[i][7] * Units.EMU_PER_PIXEL);
                    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                    drawing.createPicture(anchor, pictureIdx);
                }else{
                    System.out.println("Image not found: " + imagePath);
                }
            }
            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOutputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void insertImageIntoExcelCenter(String excelFilePath, int sheetIndex, String[] imagesPath,int[][] imagePoints,double[] scaleFactors) {
        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(excelFilePath))) {
            // 打开Excel文件
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            // 获取画图管理器
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            CreationHelper helper = workbook.getCreationHelper();

            for (int i = 0; i < imagesPath.length; i++) {
                String imagePath = imagesPath[i];
                if (new File(imagePath).exists()){
                    InputStream inputStream = new FileInputStream(imagePath);
                    byte[] bytes = IOUtils.toByteArray(inputStream);
                    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                    inputStream.close();

                    BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imagePath));
                    //图片的真实长宽
                    int pictureWidth = bufferedImage.getWidth();
                    int pictureHeight = bufferedImage.getHeight();
                    // 获取图片起始单元格
                    int col1 = imagePoints[i][0];
                    int row1 = imagePoints[i][1];
                    int col2 = imagePoints[i][2];
                    int row2 = imagePoints[i][3];

                    // 计算单元格的总宽度
                    float cellWidth = 0;
                    for (int col = col1; col <= col2; col++) {
                        cellWidth += sheet.getColumnWidthInPixels(col);
                    }
                    // 计算单元格的总高度
                    float cellHeight = 0;
                    for (int r = row1; r <= row2; r++) {
                        Row row = sheet.getRow(r);
//                        if (row == null) {
//                            row = sheet.createRow(r); // 创建一个新的空行
//                        }
//                        cellHeight += row.getHeightInPoints() / 72 * 96; // 转换为像素
                        if (row != null) {  // 检查 row 是否为 null
                            cellHeight += row.getHeightInPoints() / 72 * 96; // 转换为像素
                        } else {
                            System.out.println("insertImageIntoExcelCenter:Row " + r + " is null.");
                            // 如果该行为空，使用默认高度
                            float defaultHeight = sheet.getDefaultRowHeightInPoints(); // 获取默认行高
                            cellHeight += defaultHeight / 72 * 96; // 转换为像素
                        }
                    }
                    // 计算缩放比例，保持图片的宽高比一致
                    //double scale = Math.max(cellWidth / pictureWidth, cellHeight / pictureHeight);
                    double scale = Math.min(cellWidth / pictureWidth, cellHeight / pictureHeight);

                    scale = scale * scaleFactors[i];
                    // 计算缩放后的图片宽高
                    int newWidth = (int) (pictureWidth * scale);
                    int newHeight = (int) (pictureHeight * scale);

                    // 计算偏移量，使图片居中
                    int dx1 = (int) ((cellWidth - newWidth) / 2);
                    int dy1 = (int) ((cellHeight - newHeight) / 2);
                    int dx2 = dx1 + newWidth;
                    int dy2 = dy1 + newHeight;
                    // 设置锚点并插入图片
                    ClientAnchor anchor = helper.createClientAnchor();
                    anchor.setCol1(col1);
                    anchor.setRow1(row1);
                    anchor.setDx1(dx1 * Units.EMU_PER_PIXEL);
                    anchor.setDy1(dy1 * Units.EMU_PER_PIXEL);
                    anchor.setDx2(dx2 * Units.EMU_PER_PIXEL);
                    anchor.setDy2(dy2 * Units.EMU_PER_PIXEL);
                    anchor.setCol2(col2);
                    anchor.setRow2(row2);
                    //anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                    Picture picture = drawing.createPicture(anchor, pictureIdx);
                }else{
                    System.out.println("Image not found: " + imagePath);
                }
            }
            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOutputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public double insertImages(String localsyPath, int sheetIndex,Map<String, String> imageMap) {
        double scale = 1.0;
        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(localsyPath))) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            CreationHelper helper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        // 检查单元格内容是否与关键字匹配
                        if (imageMap.containsKey(cellValue)) {
                            String imagePath = imageMap.get(cellValue);
                            // 清除单元格内容
                            cell.setCellValue("");
                            // 获取合并单元格的范围
                            CellRangeAddress mergedRegion = getMergedRegion(sheet, row.getRowNum(), cell.getColumnIndex());
                            if ("(检验检测专用章)".equals(cellValue)){
                                mergedRegion.setFirstRow(mergedRegion.getFirstRow() - 4);
                                mergedRegion.setLastRow(mergedRegion.getLastRow() + 1);
                                mergedRegion.setLastColumn(mergedRegion.getLastColumn());
                            }
                            // 计算合并单元格的总宽度
                            float cellWidth = 0;
                            for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
                                cellWidth += sheet.getColumnWidthInPixels(col);
                            }
                            // 计算合并单元格的总高度
                            float cellHeight = 0;
                            for (int r = mergedRegion.getFirstRow(); r <= mergedRegion.getLastRow(); r++) {
                                Row row1 = sheet.getRow(r);
//                                if (row1 == null) {
//                                    row1 = sheet.createRow(r); // 创建一个新的空行
//                                }
//                                cellHeight += row1.getHeightInPoints() / 72 * 96; // 转换为像素
                                if (row1 != null) {  // 检查 row 是否为 null
                                    cellHeight += row1.getHeightInPoints() / 72 * 96; // 转换为像素
                                } else {
                                    System.out.println("insertImages:Row " + r + " is null.");
                                    // 如果该行为空，使用默认高度
                                    float defaultHeight = sheet.getDefaultRowHeightInPoints(); // 获取默认行高
                                    cellHeight += defaultHeight / 72 * 96; // 转换为像素
                                }
                            }
                            try (InputStream is = new FileInputStream(imagePath)){
                                byte[] bytes = org.apache.commons.io.IOUtils.toByteArray(is);
                                int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                                // 获取图片的原始宽高
                                BufferedImage image = ImageIO.read(new FileInputStream(imagePath));
                                int pictureWidth = image.getWidth();
                                int pictureHeight = image.getHeight();

                                // 计算缩放比例，保持图片的宽高比一致
                                scale = Math.min(cellWidth / pictureWidth, cellHeight / pictureHeight);
                                // 计算缩放后的图片宽高
                                int newWidth = (int) (pictureWidth * scale);
                                int newHeight = (int) (pictureHeight * scale);

                                // 计算偏移量，使图片居中
                                int dx1 = (int) ((cellWidth - newWidth) / 2);
                                int dy1 = (int) ((cellHeight - newHeight) / 2);
                                //int dx2 = dx1 + newWidth;
                                //int dy2 = dy1 + newHeight;
                                // 设置锚点并插入图片
                                ClientAnchor anchor = helper.createClientAnchor();
                                anchor.setCol1(mergedRegion.getFirstColumn());
                                anchor.setRow1(mergedRegion.getFirstRow());
                                anchor.setDx1(dx1 * Units.EMU_PER_PIXEL);
                                anchor.setDy1(dy1 * Units.EMU_PER_PIXEL);
                                //anchor.setDx2(dx2 * Units.EMU_PER_PIXEL);
                                //anchor.setDy2(dy2 * Units.EMU_PER_PIXEL);
                                //anchor.setCol2(mergedRegion.getLastColumn());
                                //anchor.setRow2(mergedRegion.getLastRow());

                                anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                                Picture picture = drawing.createPicture(anchor,pictureIdx);
                                picture.resize(scale);
                                //固定好了dx1，dy1，dx2，dy2，然后就可以调整宽高了，不需要使用resize方法缩放
                                ////设置单元格居中
                                //CellStyle style = workbook.createCellStyle();
                                //style.setAlignment(HorizontalAlignment.CENTER);
                                //style.setVerticalAlignment(VerticalAlignment.CENTER);
                                //cell.setCellStyle(style);
                            }catch (IOException e) {
                                throw new RuntimeException("Error processing image: " + imagePath, e);
                            }
                        }
                    }
                }
            }
            try (FileOutputStream fileOutputStream = new FileOutputStream(localsyPath)) {
                workbook.write(fileOutputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return scale;
    }

    /**
     * 获取合并单元格的范围
     *
     * @param sheet  工作表
     * @param row    单元格行号
     * @param column 单元格列号
     * @return 合并单元格的范围
     */
    public CellRangeAddress getMergedRegion(Sheet sheet, int row, int column) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(row, column)) {
                return merged;
            }
        }
        return new CellRangeAddress(row, row, column, column);
    }
    public boolean removeExistingImage(String filename,String pictureName) {
        boolean isRemoved = false;
        try(Workbook workbook = WorkbookFactory.create(new FileInputStream(filename))){
            Sheet sheet = workbook.getSheetAt(0);

            Drawing drawing = sheet.getDrawingPatriarch();

            XSSFPicture xssfPictureToDelete = null;
            if (drawing instanceof XSSFDrawing) {
                for (XSSFShape shape : ((XSSFDrawing)drawing).getShapes()) {
                    if (shape instanceof XSSFPicture) {
                        XSSFPicture xssfPicture = (XSSFPicture)shape;
                        String shapename = xssfPicture.getShapeName();
                        int row = xssfPicture.getClientAnchor().getRow1();
                        int col = xssfPicture.getClientAnchor().getCol1();
                        System.out.println("Picture " + "" + " with Shapename: " + shapename + " is located row: " + row + ", col: " + col);
                        if (pictureName.equals(shapename)) xssfPictureToDelete = xssfPicture;
                    }
                }
            }
            if (xssfPictureToDelete != null){
                isRemoved = deleteEmbeddedXSSFPicture(xssfPictureToDelete);
            }
            if (xssfPictureToDelete != null){
                isRemoved = deleteCTAnchor(xssfPictureToDelete);
            }
            //HSSFPicture hssfPictureToDelete = null;
            //if (drawing instanceof HSSFPatriarch) {
            //    for (HSSFShape shape : ((HSSFPatriarch)drawing).getChildren()) {
            //        if (shape instanceof HSSFPicture) {
            //            HSSFPicture hssfPicture = (HSSFPicture)shape;
            //            int picIndex = hssfPicture.getPictureIndex();
            //            String shapename = hssfPicture.getShapeName().trim();
            //            int row = hssfPicture.getClientAnchor().getRow1();
            //            int col = hssfPicture.getClientAnchor().getCol1();
            //            System.out.println("Picture " + picIndex + " with Shapename: " + shapename + " is located row: " + row + ", col: " + col);
            //
            //            if ("Image 2".equals(shapename)) hssfPictureToDelete = hssfPicture;
            //
            //        }
            //    }
            //}
            //if (hssfPictureToDelete != null) deleteHSSFShape(hssfPictureToDelete);
            if(isRemoved){
                try (FileOutputStream fileOutputStream = new FileOutputStream(filename)) {
                    workbook.write(fileOutputStream);
                }
            }
            return isRemoved;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public boolean deleteCTAnchor(XSSFPicture xssfPicture) {
        boolean isDeleted = false;
        XSSFDrawing drawing = xssfPicture.getDrawing();
        XmlCursor cursor = xssfPicture.getCTPicture().newCursor();
        cursor.toParent();
        if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor) {
            for (int i = 0; i < drawing.getCTDrawing().getTwoCellAnchorList().size(); i++) {
                if (cursor.getObject().equals(drawing.getCTDrawing().getTwoCellAnchorArray(i))) {
                    drawing.getCTDrawing().removeTwoCellAnchor(i);
                    isDeleted = true;
                    System.out.println("TwoCellAnchor for picture " + xssfPicture + " was deleted.");
                }
            }
        } else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor) {
            for (int i = 0; i < drawing.getCTDrawing().getOneCellAnchorList().size(); i++) {
                if (cursor.getObject().equals(drawing.getCTDrawing().getOneCellAnchorArray(i))) {
                    drawing.getCTDrawing().removeOneCellAnchor(i);
                    isDeleted = true;
                    System.out.println("OneCellAnchor for picture " + xssfPicture + " was deleted.");
                }
            }
        } else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor) {
            for (int i = 0; i < drawing.getCTDrawing().getAbsoluteAnchorList().size(); i++) {
                if (cursor.getObject().equals(drawing.getCTDrawing().getAbsoluteAnchorArray(i))) {
                    drawing.getCTDrawing().removeAbsoluteAnchor(i);
                    isDeleted = true;
                    System.out.println("AbsoluteAnchor for picture " + xssfPicture + " was deleted.");
                }
            }
        }
        return isDeleted;
    }

    public boolean deleteEmbeddedXSSFPicture(XSSFPicture xssfPicture) {
        boolean isDeleted = false;
        if (xssfPicture.getCTPicture().getBlipFill() != null) {
            if (xssfPicture.getCTPicture().getBlipFill().getBlip() != null) {
                if (xssfPicture.getCTPicture().getBlipFill().getBlip().getEmbed() != null) {
                    String rId = xssfPicture.getCTPicture().getBlipFill().getBlip().getEmbed();
                    XSSFDrawing drawing = xssfPicture.getDrawing();
                    drawing.getPackagePart().removeRelationship(rId);
                    drawing.getPackagePart().getPackage().deletePartRecursive(drawing.getRelationById(rId).getPackagePart().getPartName());
                    isDeleted = true;
                    System.out.println("Picture " + xssfPicture + " was deleted.");
                }
            }
        }
        return isDeleted;
    }
    public void deleteHSSFShape(HSSFShape shape) {
        HSSFPatriarch drawing = shape.getPatriarch();
        drawing.removeShape(shape);
        System.out.println("Shape " + shape + " was deleted.");
    }




}
