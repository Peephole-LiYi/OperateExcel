package org.dreamstu;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * @author Administrator
 * @Methodname Main
 * @description TODO
 * @date 2023/9/28 17:59
 */
public class Main {
    public static void main(String[] args) throws Exception{
        String sourceFile = "C:\\Users\\Administrator\\Desktop\\人员信息\\交通技术学校人行道闸师生信息\\21幼儿保育1班.xlsx";
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

        for (int i = sourceRowIndex; i < sourceSheet.getLastRowNum(); i++) {

            //获取表中指定坐标的数据,要求是指定行数之后的 RowIndex
            Row sourceSheetRow = sourceSheet.getRow(i);
            Row targetSheetRow = targetSheet.getRow(i);

            //如果指定行不为空的情况
            if (sourceSheetRow != null && targetSheetRow != null) {
                for (int j = sourceCellIndex; j < sourceSheetRow.getLastCellNum(); j++) {
                    Cell sourceCell = sourceSheetRow.getCell(j);
                    Cell targetCell = targetSheetRow.createCell(j);

                    if (sourceCell != null) {
                        CellStyle newCellStyle = targetWorkbook.createCellStyle();
                        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
                        targetCell.setCellStyle(newCellStyle);

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
                            default:
                                break;
                        }
                    }
                }

            }else if (sourceSheetRow == null && targetSheetRow == null){
                // 如果当前行为空，则在目标Excel文件中添加一个空行
                if (sourceSheetRow.getLastCellNum() == 0) {
                    int targetRowNum = targetSheet.getLastRowNum() + 1;
                    targetSheet.createRow(targetRowNum);
                }
            }
        }

        FileOutputStream targetWorkbookOutput = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\处理过的\\人员信息导入.xlsx");
        targetWorkbook.write(targetWorkbookOutput);
        targetWorkbookOutput.close();
        sourceWorkbook.close();
        targetWorkbook.close();
    }
}

