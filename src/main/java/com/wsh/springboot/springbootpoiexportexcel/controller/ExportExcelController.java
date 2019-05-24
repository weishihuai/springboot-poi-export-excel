package com.wsh.springboot.springbootpoiexportexcel.controller;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description: 动态生成Excel模板控制层
 * @author: weishihuai
 * @Date: 2019/4/15 14:58
 *
 * <p>
 * 说明:
 * 创建Excel（.xls后缀）大致步骤：
 * 1、创建HSSFWorkbook对象（也就是excel文档对象）;
 * 2、通过HSSFWorkbook对象创建HSSFSheet对象（也就是excel中的sheet）;
 * 3、通过HSSFSheet对象创建HSSFROW对象（Excel行）;
 * 4、通过HSSFROW对象创建列HSSFCell(Excel列)并set值（列名）;
 * 5、接着就是设置Excel的一些样式、批注、合并单元格等操作;
 * <p>
 * 创建Excel（.xlsx后缀）大致步骤：
 * 1、创建XSSFWorkbook对象（也就是excel文档对象）;
 * 2、通过XSSFWorkbook对象创建XSSFSheet对象（也就是excel中的sheet）;
 * 3、通过XSSFSheet对象创建XSSFROW对象（Excel行）;
 * 4、通过XSSFROW对象创建列XSSFCell(Excel列)并set值（列名）;
 * 5、接着就是设置Excel的一些样式、批注、合并单元格等操作;
 * <p>
 * 特别需要注意点:
 * 1. HSSFWorkbook/HSSFSheet...以H开头的一些列类:  .xls后缀名的Excel ,并且需要设置 response.setContentType("application/vnd.ms-excel;charset=gb2312");
 * 2. XSSFWorkbook/XSSFSheet/XSSFRow ...以H开头的一些列类: .xlsx文件后缀名的Excel,并且需要设置 response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
 * 3. 关于前端调用方式: window.open(window.baseUrl + "/exportDynamicExcelTemplate", "_blank") 动态生成Excel模板;
 */
@RestController
public class ExportExcelController {

    @GetMapping("/exportDynamicExcelTemplate")
    public void exportDynamicExcelTemplate(HttpServletResponse response) {
        //创建一个工作簿(Excel的文档对象)
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

        /**
         * 1. Excel文档摘要信息(可通过右键->Excel文件->属性看见)
         */
        this.createDocumentInfo(hssfWorkbook);

        //创建一个Sheet(Excel的表单)
        HSSFSheet hssfSheet = hssfWorkbook.createSheet("学生信息导入模板表");
        //创建Excel表头(rownum:0表示表头),从0开始
        HSSFRow headerRow = hssfSheet.createRow(0);


        /**
         * 2. 设置行高、列宽、单元格样式、字体颜色、大小等
         */
        // 设置缺省列高
        hssfSheet.setDefaultRowHeightInPoints(25);
        //设置第一行(表头)的高度
        headerRow.setHeightInPoints(25);
        //设置列宽，设置第8列宽度是60个字符宽度
        hssfSheet.setColumnWidth(8, 40 * 256);
        // 设置表格默认列宽度为15个字节
        hssfSheet.setDefaultColumnWidth((short) 16);
        //格子单元样式
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        HSSFFont hssfFont = hssfWorkbook.createFont();
        //设置字体加粗
        hssfFont.setBold(true);
        //设置字体名称
        hssfFont.setFontName("华文行楷");
        //设置字体大小
        hssfFont.setFontHeightInPoints((short) 15);
        //设置下划线
        //hssfFont.setUnderline(FontFormatting.U_SINGLE);
        //设置删除线
        //hssfFont.setStrikeout(true);
        hssfCellStyle.setFont(hssfFont);
        //设置水平居中
        hssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        //设置垂直居中
        hssfCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        /**
         * 3. Excel模板表头相关信息设置
         * 导出的Excel的标题列
         * 实际工作可通过查询数据库，搞成动态表头Excel
         */
        HSSFFont hFont = hssfWorkbook.createFont();
        HSSFCellStyle hCellStyle;
        String[] titleHeaders = {"学号", "姓名", "性别", "学院", "专业", "班级", "入学日期", "学费", "学生类别", "住宿费"};
        for (int i = 0; i < titleHeaders.length; i++) {
            String header = titleHeaders[i];
            HSSFCell cell = headerRow.createCell(i);
            HSSFRichTextString text = new HSSFRichTextString(header);
            cell.setCellValue(text);
            //设置学号/姓名列一些必填标识等
            if ("学号".equals(header) || "姓名".equals(header)) {
                //设置字体样式等
                hCellStyle = this.createCellStyle(hFont, hssfWorkbook);
                cell.setCellStyle(hCellStyle);
            } else {
                cell.setCellStyle(hssfCellStyle);
            }
        }

        //构造两条示例数据
        List<Map<String, Object>> exampleDataList = getExampleData();
        for (int i = 0; i < exampleDataList.size(); i++) {
            Map<String, Object> map = exampleDataList.get(i);
            if (null != map && !map.isEmpty()) {
                //创建Excel一行(第二行)
                HSSFRow row = hssfSheet.createRow(i + 1);

                /**
                 * 4. 创建表格列、设置单元格数据等
                 */
                //创建一列(column:0 --> 第一列),从0开始
                row.createCell(0).setCellValue(null != map.get("xh") ? map.get("xh").toString() : "");
                //setCellValue: 设置单元格内容
                row.createCell(1).setCellValue(null != map.get("xm") ? map.get("xm").toString() : "");
                row.createCell(2).setCellValue(null != map.get("xb") ? map.get("xb").toString() : "");
                row.createCell(3).setCellValue(null != map.get("xy") ? map.get("xy").toString() : "");
                row.createCell(4).setCellValue(null != map.get("zy") ? map.get("zy").toString() : "");
                HSSFCell bjCell = row.createCell(5);
                bjCell.setCellValue(null != map.get("bj") ? map.get("bj").toString() : "");
                if ("201324131212".equals(null != map.get("xh") ? map.get("xh").toString() : "")) {
                    /**
                     * 5. 给某个单元格创建批注信息
                     */
                    HSSFComment hssfComment = this.createPatriarch(hssfSheet);
                    //把批注赋值给单元格
                    bjCell.setCellComment(hssfComment);
                }

                /**
                 * 6. 处理特定的数据格式(如日期、金额、小数等)
                 */
                //设置日期数据格式
                HSSFCell rxrqCell = row.createCell(6);
                rxrqCell.setCellValue(new Date());
                HSSFCellStyle style = hssfWorkbook.createCellStyle();
                //格式化日期格式数据
                style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
                rxrqCell.setCellStyle(style);

                //设置保留2位小数(四舍五入)
                HSSFCell xfCell = row.createCell(7);
                xfCell.setCellValue(20000.555555);
                style = hssfWorkbook.createCellStyle();
                style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                xfCell.setCellStyle(style);

                //货币格式/自定义格式
                HSSFCell hbCell = row.createCell(9);
                hbCell.setCellValue(1200.0000);
                style = hssfWorkbook.createCellStyle();
                style.setDataFormat(hssfWorkbook.createDataFormat().getFormat("￥#,##0"));
                hbCell.setCellStyle(style);
            }
        }

        /**
         * 7. 合并单元格(注意从0开始)
         * 说明:
         *  firstRow 合并区域中第一个单元格的行号
         *  lastRow  合并区域中最后一个单元格的行号
         *  firstCol 合并区域中第一个单元格的列号
         *  lastCol  合并区域中最后一个单元格的列号
         */
        //合并行 ----> 对应Excel表格的 第4行到第5行、第1列到第7列
        CellRangeAddress cellRangeAddress = new CellRangeAddress(3, 4, 0, 6);
        hssfSheet.addMergedRegion(cellRangeAddress);
        //合并行 ----> 对应Excel表格的 第1行到第6行、第11列到第12列
        cellRangeAddress = new CellRangeAddress(0, 5, 10, 11);
        hssfSheet.addMergedRegion(cellRangeAddress);

        /**
         * 8. 动态生成下拉式菜单
         * 说明: 必要的时候可通过数据库查询动态替换下拉数据
         */
        HSSFDataValidation hssfDataValidation = this.createDynamicSelectOptions();
        //对sheet页生效
        hssfSheet.addValidationData(hssfDataValidation);

        //导出的Excel模板文件名称
        DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
        // 需要指定文件名的编码方式为ISO-8859-1,否则会出现文件名中的中文字符丢失的问题   //比如: _________20190415.xls
        String excelFileName = null;
        OutputStream os = null;
        try {
            excelFileName = new String(("学生信息导入模板表" + dateFormat.format(new Date()) + ".xls").getBytes("gb2312"), "ISO-8859-1");
            //清空response
            response.reset();
            //设置response的Header
            response.addHeader("Content-Disposition", "attachment;filename=" + excelFileName);
            os = new BufferedOutputStream(response.getOutputStream());
            //设置消息头内容格式,并指定编码格式
            response.setContentType("application/vnd.ms-excel;charset=gb2312");
            //将excel写入到输出流中
            hssfWorkbook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != os) {
                    os.flush();
                    os.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 设置单元格样式
     *
     * @param hssfFont 字体
     * @return
     */
    private HSSFCellStyle createCellStyle(HSSFFont hssfFont, HSSFWorkbook hssfWorkbook) {
        HSSFCellStyle hssfCellStyle2 = hssfWorkbook.createCellStyle();
        //粗体
        hssfFont.setBold(true);
        //设置字体名称
        hssfFont.setFontName("华文行楷");
        //设置字体大小
        hssfFont.setFontHeightInPoints((short) 15);
        //字体颜色
        hssfFont.setColor(IndexedColors.RED.getIndex());
        hssfCellStyle2.setFont(hssfFont);
        //水平居中
        hssfCellStyle2.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        hssfCellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置图案样式
        hssfCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //设置图案颜色
        hssfCellStyle2.setFillForegroundColor(IndexedColors.YELLOW.index);
        return hssfCellStyle2;
    }

    /**
     * 动态构造下拉数据
     */
    private HSSFDataValidation createDynamicSelectOptions() {
        //在第1行第9列(学生类别列)，插入下拉选择框,这里将lastRow设置为很大的一个数就是为了使这一列的值只能通过下拉框选择
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(0, 65535, 8, 8);
        //生成下拉框内容
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(new String[]{"本科生", "专科生", "其他类型"});
        //绑定下拉框和作用区域
        HSSFDataValidation hssfDataValidation = new HSSFDataValidation(cellRangeAddressList, constraint);
        //单元格输入提示信息
        hssfDataValidation.createPromptBox("输入提示", "请从下拉列表中选择学生类别");
        //是否显示提示信息
        hssfDataValidation.setShowPromptBox(true);
        return hssfDataValidation;
    }

    /**
     * 创建批注
     *
     * @param hssfSheet Sheet表单
     * @return 批注信息
     */
    private HSSFComment createPatriarch(HSSFSheet hssfSheet) {
        //创建批注
        HSSFPatriarch hssfPatriarch = hssfSheet.createDrawingPatriarch();
        //创建批注位置
        HSSFClientAnchor hssfClientAnchor = hssfPatriarch.createAnchor(0, 0, 0, 0, 6, 1, 7, 3);
        //批注
        HSSFComment hssfComment = hssfPatriarch.createCellComment(hssfClientAnchor);
        //设置批注内容
        hssfComment.setString(new HSSFRichTextString("这是一个班级单元格！"));
        //设置批注作者
        hssfComment.setAuthor("weixiaohuai");
        //设置批注默认显示
        hssfComment.setVisible(true);
        return hssfComment;
    }

    /**
     * 创建文档信息
     *
     * @param hssfWorkbook Excel工作簿
     */
    private void createDocumentInfo(HSSFWorkbook hssfWorkbook) {
        //创建文档信息
        hssfWorkbook.createInformationProperties();
        //摘要信息
        DocumentSummaryInformation information = hssfWorkbook.getDocumentSummaryInformation();
        //设置类别
        information.setCategory("学生导入模板");
        //设置文档管理者名称
        information.setManager("weixiaohuai");
        //设置公司
        information.setCompany("gzly");
        SummaryInformation summaryInformation = hssfWorkbook.getSummaryInformation();
        //作者
        summaryInformation.setAuthor("weixiaohuai");
        //备注
        summaryInformation.setComments("动态生成学生导入模板");
        //主题
        summaryInformation.setSubject("学生模板");
        //标题
        summaryInformation.setTitle("学生导入模板");
    }

    /**
     * 获取Excel模板示例数据
     */
    private List<Map<String, Object>> getExampleData() {
        List<Map<String, Object>> mapList = new ArrayList<>();
        //初始化两条示例数据
        Map<String, Object> map = new HashMap<>();
        map.put("xh", "201324131147");
        map.put("xm", "张三");
        map.put("xb", "男");
        map.put("xy", "计算机学院");
        map.put("zy", "软件工程");
        map.put("bj", "15科技一班");
        map.put("xf", 20000.00);
        mapList.add(map);
        Map<String, Object> map2 = new HashMap<>();
        map2.put("xh", "201324131212");
        map2.put("xm", "李四");
        map2.put("xb", "男");
        map2.put("xy", "数学学院");
        map2.put("zy", "高等数学");
        map2.put("bj", "16高数一班");
        map.put("xf", 20000.00);
        mapList.add(map2);
        return mapList;
    }

}
