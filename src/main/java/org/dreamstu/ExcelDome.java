package org.dreamstu;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


import java.io.*;
import java.util.Scanner;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import java.util.zip.Deflater;
import java.util.zip.ZipOutputStream;


/**
 * @author Administrator
 * @Methodname ExcelDome
 * @description 方法入口
 * @date 2023/10/12 21:02
 */
public class ExcelDome {

    public static Scanner scanner = new  Scanner(System.in);
    public static final Logger logger = Logger.getLogger(ExcelDome.class.getName());
    public static void main(String[] args) {

        try {
            //在读取文件前设置文件压缩率为-1.0d，以防出现zip炸弹
            ZipSecureFile.setMinInflateRatio(-1.0d);

            FileHandler fileHandler = new FileHandler("log.txt", true);
            SimpleFormatter formatter = new SimpleFormatter();
            fileHandler.setFormatter(formatter);
            logger.addHandler(fileHandler);

            while (true) {
                System.out.println("输入1.进入程序");
                System.out.println("输入2.退出程序");
                String commod = scanner.next();
                switch (commod){
                    case "1":
                        System.out.println("请输入源文件地址:");
                        String sourceFile = scanner.next();
                        System.out.println("请输入目标文件地址:");
                        String targetFile = scanner.next();
                        System.out.println("目前替换的源列为：0.1.2.3.5.6.7  对应目标文件列为：0.1.2.7.9.10.11 从第三行开始");

                        int sourceSheetIndex = 0;
                        int sourceRowIndex = 3;
                        int sourceCellIndex = 0;
                        int targetSheetIndex = 0;
                        int targetRowIndex = 3;
                        int targetCellIndex = 0;

                        copyData(sourceFile,
                                targetFile,
                                sourceSheetIndex,
                                sourceRowIndex,
                                sourceCellIndex,
                                targetSheetIndex,
                                targetRowIndex,
                                targetCellIndex);
                        break;
                    case "2":
                        logger.log(Level.INFO, "程序退出！");
                        return;
                    default:
                        System.out.println("输入有误!请重新输入.");
                        logger.log(Level.INFO, "输入有误！");
                        break;
                }
            }
        } catch (Exception e) {
            String errorMessage = e.getMessage();
            logger.log(Level.SEVERE, errorMessage);
        }

    }

    /**
     * @description: copy表格方法
     * @author: peephole
     * @date: 2023/10/12 21:59
     * @param:
     * @return: true/false
     **/
    private static void copyData(String sourceFile, String targetFile, int sourceSheetIndex, int sourceRowIndex, int sourceCellIndex, int targetSheetIndex, int targetRowIndex, int targetCellIndex) throws Exception {

        InputStream sourceStream = new FileInputStream(sourceFile);
        InputStream targetStream = new FileInputStream(targetFile);



        //创建一个工作簿
        Workbook sourceWorkbook = new XSSFWorkbook(sourceStream);
        Workbook targetWorkbook = new XSSFWorkbook(targetStream);
        //读取指定索引的表
        Sheet sourceSheet = sourceWorkbook.getSheetAt(sourceSheetIndex);
        Sheet targetSheet = targetWorkbook.getSheetAt(targetSheetIndex);
        //读取指定表下的行

        for (int i = sourceRowIndex; i <= sourceSheet.getLastRowNum(); i++) {

            //获取表中指定坐标的数据,要求是指定行数之后的 RowIndex
            Row sourceSheetRow = sourceSheet.getRow(i);
            Row targetSheetRow = targetSheet.getRow(i);

//            Cell cell = sourceSheetRow.getCell(i);
//            System.out.println(cell);
            //如果指定行不为空的情况
            if (sourceSheetRow != null && targetSheetRow != null) {

                //遍历列
                for (int j = sourceCellIndex; j < sourceSheetRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceSheetRow.getCell(j);
                    Cell targetCell = targetSheetRow.createCell(j);



                    //System.out.println(sourceCell);
//                    if (sourceCell != null) {
                    if (sourceCell != null){
                        CellStyle newCellStyle = targetWorkbook.createCellStyle();
                        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
                        targetCell.setCellStyle(newCellStyle);
                    }
                    //System.out.println(sourceCell);
                    //判断j在哪行,创建目标文件对应列数据.


                    if (j == 1){

                        setRowCellData(1, targetSheetRow, sourceCell);
                    }

                    if (j == 0){
                        setRowCellData(0, targetSheetRow, sourceCell);
                        logger.log(Level.INFO, sourceCell + "已经更改");
                    }

                    if (j == 2){
                        setRowCellData(2, targetSheetRow, sourceCell);
                    }

                    if (j == 3){
                        //设置目标行的样式
                        CellStyle newCellStyle = targetWorkbook.createCellStyle();
                        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
                        targetCell.setCellStyle(newCellStyle);

                        System.out.println(sourceCell);
                        //CellType cellType = sourceCell.getCellType();
                        String cellValue = sourceCell.getStringCellValue();
//                        System.out.println(cellType);
                        Cell targetSheetRowCell = targetSheetRow.createCell(8);
                        targetSheetRowCell.setCellValue(cellValue);
                        //setRowCellData(7, targetSheetRow, sourceCell);

                    }

                    if (j == 5){
                        continue;
                    }

                    if (j == 4){
                        if (sourceCell == null){
                            System.out.println("错误!,指定列为空");
                            break;
                        }
                        String stringCellValue = sourceCell.getStringCellValue();
                        //System.out.println(stringCellValue);
                        Cell targetSheetRowCell9 = targetSheetRow.createCell(9);

                        if (stringCellValue.equals("是")){
                            targetSheetRowCell9.setCellValue("住宿");
                        }else if (stringCellValue.equals("否")){
                            targetSheetRowCell9.setCellValue("走读");
                        }else {
                            logger.log(Level.WARNING, sourceCell + "的元数据非是/否!!!!!");

                        }
                        //System.out.println("当前值为:" + sourceCell.getStringCellValue());
                    }

                    if (j == 6){
                        setRowCellData(10, targetSheetRow, sourceCell);
                    }

                    if (j == 7){
                        setRowCellData(11, targetSheetRow, sourceCell);
                    }
                }
            }
        }

        FileOutputStream targetWorkbookOutput = new FileOutputStream(targetFile);
        targetWorkbook.write(targetWorkbookOutput);
        targetWorkbookOutput.close();
        sourceWorkbook.close();
        targetWorkbook.close();
    }

    /**
     *
     * @author peephole
     * @date 2023/10/12 21:46
     * @methodName setRowCellData
     * @description 根据传入单元格判断当前数据为什么类型然后赋值给当前单元格
     * @return 无返回值
     */

    private static void setRowCellData(int i, Row targetSheetRow, Cell sourceCell) {
        if (sourceCell == null){
            logger.log(Level.SEVERE,"错误!,指定列为空");
        }
        Cell targetSheetRowCell = targetSheetRow.createCell(i);
        switch (sourceCell.getCellType()) {
            case STRING:
                targetSheetRowCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                targetSheetRowCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BOOLEAN:
                targetSheetRowCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetSheetRowCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }
}