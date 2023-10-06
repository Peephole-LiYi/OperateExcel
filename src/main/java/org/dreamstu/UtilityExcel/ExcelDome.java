package org.dreamstu.dome;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


import java.io.*;


/**
 * @author Administrator
 * @Methodname ExcelDome
 * @description TODO
 * @date 2023/9/27 22:33
 */
public class ExcelDome {
    @Test
    public void ExcelDomeMethd() throws Exception {
        String sourceFile = "C:\\Users\\Administrator\\Desktop\\人员信息\\交通技术学校人行道闸师生信息\\21幼儿保育2班.xlsx";
        String targetFile = "C:\\Users\\Administrator\\Desktop\\处理过的\\人员信息导入.xlsx";

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

    }

    /**
     * @description: copy表格方法
     * @author: peephole
     * @date:
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

            //如果指定行不为空的情况
            if (sourceSheetRow != null && targetSheetRow != null) {

                //遍历列
                for (int j = sourceCellIndex; j < sourceSheetRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceSheetRow.getCell(j);
                    Cell targetCell = targetSheetRow.createCell(j);

                    System.out.println(sourceCell);
//                    if (sourceCell != null) {
                    if (sourceCell != null){
                        CellStyle newCellStyle = targetWorkbook.createCellStyle();
                        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
                        targetCell.setCellStyle(newCellStyle);
                    }


                        if (j == 1){
                            Cell targetSheetRowCell1 = targetSheetRow.createCell(1);
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetSheetRowCell1.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetSheetRowCell1.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetSheetRowCell1.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetSheetRowCell1.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    break;
                            }
                        }

                        if (j == 0){
                            Cell targetSheetRowCell0 = targetSheetRow.createCell(0);
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetSheetRowCell0.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetSheetRowCell0.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetSheetRowCell0.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetSheetRowCell0.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    break;
                            }
                        }

                        if (j == 2){
                            Cell targetSheetRowCell2 = targetSheetRow.createCell(2);
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetSheetRowCell2.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetSheetRowCell2.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetSheetRowCell2.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetSheetRowCell2.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    break;
                            }
                        }

                        if (j == 3){
                            Cell targetSheetRowCell7 = targetSheetRow.createCell(7);
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetSheetRowCell7.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetSheetRowCell7.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetSheetRowCell7.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetSheetRowCell7.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    break;
                            }
                        }

                        if (j == 4){
                           continue;
                        }

                        if (j == 5){
                            String stringCellValue = sourceCell.getStringCellValue();
                            Cell targetSheetRowCell9 = targetSheetRow.createCell(9);
                            if (stringCellValue.equals("是")){
                                targetSheetRowCell9.setCellValue("住宿");
                            }else if (stringCellValue.equals("否")){
                                targetSheetRowCell9.setCellValue("走读");
                            }else {
                                System.out.println("元数据非是/否!!!!!");

                            }
                            //System.out.println("当前值为:" + sourceCell.getStringCellValue());
                        }

                        if (j == 6){
                            Cell targetSheetRowCell10 = targetSheetRow.createCell(10);
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetSheetRowCell10.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetSheetRowCell10.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetSheetRowCell10.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetSheetRowCell10.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    break;
                            }
                        }

                        if (j == 7){
                            Cell targetSheetRowCell11 = targetSheetRow.createCell(11);
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetSheetRowCell11.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetSheetRowCell11.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetSheetRowCell11.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetSheetRowCell11.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    break;
                            }
                        }


//                        targetCell.setCellValue(sourceCell.getStringCellValue());


//                        switch (sourceCell.getCellType()) {
//                            case STRING:
//                                targetCell.setCellValue(sourceCell.getStringCellValue());
//                                break;
//                            case NUMERIC:
//                                targetCell.setCellValue(sourceCell.getNumericCellValue());
//                                break;
//                            case BOOLEAN:
//                                targetCell.setCellValue(sourceCell.getBooleanCellValue());
//                                break;
//                            case FORMULA:
//                                targetCell.setCellFormula(sourceCell.getCellFormula());
//                                break;
//                            default:
//                                break;
//                        }
                    }
                }

//            }else if (sourceSheetRow == null && targetSheetRow == null){
//                // 如果当前行为空，则在目标Excel文件中添加一个空行
//                if (sourceSheetRow.getLastCellNum() == 0) {
//                    int targetRowNum = targetSheet.getLastRowNum() + 1;
//                    targetSheet.createRow(targetRowNum);
//                }
//            }
        }

        FileOutputStream targetWorkbookOutput = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\处理过的\\人员信息导入.xlsx");
        targetWorkbook.write(targetWorkbookOutput);
        targetWorkbookOutput.close();
        sourceWorkbook.close();
        targetWorkbook.close();
    }
}
