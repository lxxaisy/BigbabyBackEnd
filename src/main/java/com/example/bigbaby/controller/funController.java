package com.example.bigbaby.controller;

import javax.servlet.http.HttpServletResponse;

import com.example.bigbaby.entity.FileInfo;
import com.example.bigbaby.entity.FinancialExcelInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * token 控制
 *
 * @author easycloud
 */
@RestController
public class funController {
    // 当前伝票No.
    String currentIndex = "";
    // 当前凭证号生成规则优先级
    Integer weight = 10;
    // 凭证类型编码
    String typeCode = "";
    // 凭证号
    String code = "";
    // 当前相同伝票No.的数据组
    List<Row> rows = new ArrayList<Row>();

    // 11/13 问题对应 工资情况下 贷方辅助科目对应信息
    Map<String, String> creditMinorNameJPInfo = new HashMap<>() {
        {
            put("マンボー会", "224102");
            put("雇用保険", "224104");
            put("賃料（社宅）", "660264");
            put("社会保険", "224104");
            put("源泉所得税", "222103");
            put("住民税", "222103");
            put("雑費", "660231");
        }
    };

    // 11/07 660201 500101 660236
    // 生成对应科目的贷方数据 221101
    Map<String, String> debitCodeForAddCreditInfo = new HashMap<>() {
        {
            put("660201", "221101");
            put("500101", "221101");
            put("660236", "221101");
        }
    };

    // 660203 500103
    // 生成对应科目的贷方数据 221102
    Map<String, String> debitCodeForAddCreditInfo2 = new HashMap<>() {
        {
            put("660203", "221102");
            put("500103", "221102");
        }
    };

    // 需要生成辅助科目 对应科目一览
    Map<String, String> codeListForMinorCode = new HashMap<>() {
        {
            put("2202", "2202");
            put("100502", "100502");
            put("112203", "112203");
            put("112213", "112213");
            put("500130", "500130");
            put("500131", "500131");
            put("600103", "600103");
            put("660206", "0");
            put("660210", "0");
            put("660231", "0");
            put("660232", "0");
            put("660234", "0");
            put("660238", "0");
            put("660239", "0");
            put("660240", "0");
            put("660241", "0");
            put("660264", "0");
            put("660266", "0");
            put("660268", "0");
            put("660236", "0");
            put("660201", "0");
            put("660203", "0");
            put("660301", "660301");
            put("122103", "122103");
            put("224101", "224101");
            put("224102", "224102");
            put("224104", "224104");
            put("500101", "500101");
            put("500103", "500103");
        }
    };

    // 辅助科目 银行科目对应
    Map<String, String> minorCodeForBank = new HashMap<>() {
        {
            put("みずほ銀行兜町支店", "2226709:银行账户");
            put("東日本銀行大崎支店", "325798:银行账户");
            put("東京三菱UFJ銀行三宮", "5462099:银行账户");
            put("交通銀行", "61553290024742:银行账户");
            put("三井住友銀行横浜", "7032082:银行账户");
        }
    };

    // 分录合并
    // 计提工资分录
    Integer salaryGroup;
    // 支付工资分录
    Integer paySalaryGroup;
    // 代扣代缴
    Integer payForOtherGroup;
    // 计提保险
    Integer insuranceGroup;

    @PostMapping("excelReport")
    public void excelReport(@RequestBody FileInfo fileInfo, HttpServletResponse response) throws IOException {
        ClassPathResource cpr = new ClassPathResource("/templates/" + "ri-template.xlsx");
        if (fileInfo.getType().equals("zh")) {
            cpr = new ClassPathResource("/templates/" + "zh-template.xlsx");
        }
        InputStream is = cpr.getInputStream();
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        Row rowB2 = sheet.getRow(1);
        Cell cellB2 = rowB2.getCell(1);
        cellB2.setCellValue(fileInfo.getOrderNo());

        Row rowD21 = sheet.getRow(20);
        Cell cellD21 = rowD21.getCell(3);
        cellD21.setCellValue(fileInfo.getPoNo());

        Row rowO24 = sheet.getRow(23);
        Cell cellO24 = rowO24.getCell(14);
        cellO24.setCellValue(fileInfo.getCount());

        Row rowD24 = sheet.getRow(23);
        Cell cellD24 = rowD24.getCell(3);
        if (fileInfo.getType().equals("zh")) {
            cellD24.setCellValue(fileInfo.getEnName());
        } else {
            cellD24.setCellValue(fileInfo.getContractName());
        }

        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        response.setCharacterEncoding("utf-8");
        response.setHeader("content-Type", "application/vnd.ms-excel");
        response.setHeader("Content-Disposition",
                "attachment;filename=\"" + URLEncoder.encode(fileInfo.getFileName(), "UTF-8") + "\"");
        workbook.write(response.getOutputStream());
    }

    @PostMapping("excelYYReport")
    public void excelYYReport(@RequestBody List<FinancialExcelInfo> infoList, HttpServletResponse response) throws IOException {
        // reset
        reset();

        ClassPathResource cpr = new ClassPathResource("/templates/" + "template.xlsx");

        InputStream is = cpr.getInputStream();
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        // 设置金额相关列为数字格式
        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat cellDf = workbook.createDataFormat(); // 此处设置数据格式
        cellStyle.setDataFormat(cellDf.getFormat("#,##0.00_ );(#,##0.00)"));

        DecimalFormat df = new DecimalFormat("###,##0.00");

        int indexRow = 2;
        for (FinancialExcelInfo info : infoList) {
            // 如果借方科目
            if (StringUtils.hasLength(info.getDebitCode())) {
                Row row = sheet.createRow(indexRow);
                indexRow++;

                // 共同信息生成
                createCommon(row, info);
                //摘要
                row.createCell(7).setCellValue(info.getAbstractInfo());
                //科目编码
                // 11/7 设计修改 借方科目是 224102 其他应付款 再加上 勘定科目列数据为未払金时 科目替换为221101 其他逻辑不变
                // 12/14 设计修改 前提要加上是否为工资的判断
                if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1") && info.getDebitCode().equals("224102") && info.getDebitCodeJP().equals("未払金")) {
                    row.createCell(8).setCellValue("221101");
                    //支付工资分录 暂时通过修改摘要形式完成
                    Cell cell7 = row.getCell(7);
                    cell7.setCellValue("支付工资分录");
                    // 11/16 生成对应凭证号
                    if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1")) {
                        // 凭证类别编码
                        row.createCell(2).setCellValue("03");
                        // 凭证号
                        row.createCell(3).setCellValue("3");
                    }
                }
                // 11/13 借方辅助科目为社会保険 借方科目替换为224104
                else if (info.getDebitMinorNameJP().equals("社会保険")) {
                    row.createCell(8).setCellValue("224104");
                    // 11/16 生成对应凭证号
                    if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1")) {
                        // 凭证类别编码
                        row.createCell(2).setCellValue("03");
                        // 凭证号
                        row.createCell(3).setCellValue("4");
                    }
                } else {
                    row.createCell(8).setCellValue(info.getDebitCode());
                }
                //原币借方金额
                Cell cell10 = row.createCell(10);
                cell10.setCellStyle(cellStyle);
                cell10.setCellValue(info.getDebitCount());
                //本币借方金额
                if (StringUtils.hasLength(info.getDebitCount())) {
                    BigDecimal debitCount = new BigDecimal(info.getDebitCount().replaceAll(",", ""));
                    BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                    Cell cell11 = row.createCell(11);
                    cell11.setCellStyle(cellStyle);
                    cell11.setCellValue(df.format(debitCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                }

                createCodeIndex(info, row);

                // 数据源会添加一列 是否为工资 当为工资的时候 如果借方科目是
                // 1.660201 2.500101 3.660236
                // 需要额外生成一条贷方科目为221101的数据 金额和借方相同(税前金额)
                if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1") && debitCodeForAddCreditInfo.containsKey(info.getDebitCode())) {
                    // 11/16 生成对应凭证号
                    // 凭证类别编码
                    row.createCell(2).setCellValue("06");
                    // 凭证号
                    row.createCell(3).setCellValue("4");

                    // 计提工资分录 合并分录 生成一条
                    // 第一次进行生成
                    if (salaryGroup == null) {
                        salaryGroup = indexRow;
                        indexRow++;

                        Row rowTmp = sheet.createRow(salaryGroup);
                        // 共同信息生成
                        createCommon(rowTmp, info);
                        // 11/16 生成对应凭证号
                        // 凭证类别编码
                        rowTmp.createCell(2).setCellValue("06");
                        // 凭证号
                        rowTmp.createCell(3).setCellValue("4");
                        //摘要
                        rowTmp.createCell(7).setCellValue("计提工资分录");
                        //科目编码
                        rowTmp.createCell(8).setCellValue(debitCodeForAddCreditInfo.get(info.getDebitCode()));
                        //原币贷方金额
                        Cell cell18 = rowTmp.createCell(18);
                        cell18.setCellStyle(cellStyle);
                        cell18.setCellValue(info.getDebitCountWithTax());
                        //本币贷方金额
                        if (StringUtils.hasLength(info.getDebitCountWithTax())) {
                            BigDecimal creditCount = new BigDecimal(info.getDebitCountWithTax().replaceAll(",", ""));
                            BigDecimal exchange = new BigDecimal(info.getExchangeRate());
                            Cell cell19 = rowTmp.createCell(19);
                            cell19.setCellStyle(cellStyle);
                            cell19.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                        }
                    } else {
                        Row salaryGroupRow = sheet.getRow(salaryGroup);
                        //原币贷方金额
                        Cell cell18 = salaryGroupRow.getCell(18);
                        BigDecimal cell18Value = new BigDecimal(cell18.getStringCellValue().replaceAll(",", ""));
                        BigDecimal current18Value = new BigDecimal(info.getDebitCountWithTax().replaceAll(",", ""));
                        cell18.setCellValue(df.format(cell18Value.add(current18Value).setScale(2, RoundingMode.HALF_UP)));

                        //本币贷方金额
                        if (StringUtils.hasLength(info.getDebitCountWithTax())) {
                            BigDecimal creditCount = cell18Value.add(current18Value).setScale(2, RoundingMode.HALF_UP);
                            BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                            Cell cell19 = salaryGroupRow.getCell(19);
                            cell19.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                        }
                    }
                }

                // 4.660203 5.500103
                // 需要额外生成一条贷方科目为221101的数据 金额和借方相同(税前金额)
                if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1") && debitCodeForAddCreditInfo2.containsKey(info.getDebitCode())) {
                    // 11/16 生成对应凭证号
                    // 凭证类别编码
                    row.createCell(2).setCellValue("06");
                    // 凭证号
                    row.createCell(3).setCellValue("6");

                    // 支付工资分录 合并分录 生成一条
                    if (insuranceGroup == null) {
                        insuranceGroup = indexRow;
                        indexRow++;

                        Row rowTmp = sheet.createRow(insuranceGroup);
                        // 共同信息生成
                        createCommon(rowTmp, info);
                        // 11/16 生成对应凭证号
                        // 凭证类别编码
                        rowTmp.createCell(2).setCellValue("06");
                        // 凭证号
                        rowTmp.createCell(3).setCellValue("6");
                        //摘要
                        rowTmp.createCell(7).setCellValue("计提保险分录");
                        //科目编码
                        rowTmp.createCell(8).setCellValue(debitCodeForAddCreditInfo2.get(info.getDebitCode()));
                        //原币贷方金额
                        Cell cell18 = rowTmp.createCell(18);
                        cell18.setCellStyle(cellStyle);
                        cell18.setCellValue(info.getDebitCountWithTax());
                        //本币贷方金额
                        if (StringUtils.hasLength(info.getDebitCountWithTax())) {
                            BigDecimal creditCount = new BigDecimal(info.getDebitCountWithTax().replaceAll(",", ""));
                            BigDecimal exchange = new BigDecimal(info.getExchangeRate());
                            Cell cell19 = rowTmp.createCell(19);
                            cell19.setCellStyle(cellStyle);
                            cell19.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                        }
                    } else {
                        Row insuranceGroupRow = sheet.getRow(insuranceGroup);
                        //原币贷方金额
                        Cell cell18 = insuranceGroupRow.getCell(18);
                        BigDecimal cell18Value = new BigDecimal(cell18.getStringCellValue().replaceAll(",", ""));
                        BigDecimal current18Value = new BigDecimal(info.getDebitCountWithTax().replaceAll(",", ""));
                        cell18.setCellValue(df.format(cell18Value.add(current18Value).setScale(2, RoundingMode.HALF_UP)));

                        //本币贷方金额
                        if (StringUtils.hasLength(info.getDebitCountWithTax())) {
                            BigDecimal creditCount = cell18Value.add(current18Value).setScale(2, RoundingMode.HALF_UP);
                            BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                            Cell cell19 = insuranceGroupRow.getCell(19);
                            cell19.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                        }
                    }
                }

                // 辅助科目
                createMinorCode(info, row, "0");
            }

            // 如果消费税科目-借方
            if (StringUtils.hasLength(info.getConsumptionTax())) {
                Row row = sheet.createRow(indexRow);
                indexRow++;

                // 共同信息生成
                createCommon(row, info);
                row.createCell(7).setCellValue("消费税");
                //科目编码
                row.createCell(8).setCellValue(getTaxCodeByTaxInfo(info, "1"));
                //原币借方金额
                Cell cell10 = row.createCell(10);
                cell10.setCellStyle(cellStyle);
                cell10.setCellValue(info.getConsumptionTax());
                //本币借方金额
                if (StringUtils.hasLength(info.getConsumptionTax())) {
                    BigDecimal debitCount = new BigDecimal(info.getConsumptionTax().replaceAll(",", ""));
                    BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                    Cell cell11 = row.createCell(11);
                    cell11.setCellStyle(cellStyle);
                    cell11.setCellValue(df.format(debitCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                }

                createCodeIndex(info, row);
            }

            // 如果贷方科目
            if (StringUtils.hasLength(info.getCreditCode())) {
                Row row = sheet.createRow(indexRow);
                indexRow++;

                // 共同信息生成
                createCommon(row, info);

                // 11/16 生成对应凭证号
                if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1") && info.getCreditCode().equals("100502")) {
                    // 凭证类别编码
                    row.createCell(2).setCellValue("03");
                    // 凭证号
                    row.createCell(3).setCellValue("4");
                }

                //摘要
                row.createCell(7).setCellValue(info.getAbstractInfo());
                //科目编码
                // 11/7 设计修改 应收帐款科目 112203 如果在贷方科目时 转换为 112213 借方时不变
                if (info.getCreditCode().equals("112203")) {
                    row.createCell(8).setCellValue("112213");
                } else {
                    row.createCell(8).setCellValue(info.getCreditCode());
                }
                //原币贷方金额
                Cell cell18 = row.createCell(18);
                cell18.setCellStyle(cellStyle);
                cell18.setCellValue(info.getCreditCount());
                //本币贷方金额
                if (StringUtils.hasLength(info.getCreditCount())) {
                    BigDecimal creditCount = new BigDecimal(info.getCreditCount().replaceAll(",", ""));
                    BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                    Cell cell19 = row.createCell(19);
                    cell19.setCellStyle(cellStyle);
                    cell19.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                }

                createCodeIndex(info, row);

                // 数据源会添加一列 是否为工资 当为工资的时候 如果贷方科目在creditMinorNameJPInfo之中
                // 则将对应的贷方科目替换为相应value
                // 需要额外生成一条借方科目为221101的数据 金额和借方相同(税前金额)
                if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1") && creditMinorNameJPInfo.containsKey(info.getCreditMinorNameJP())) {
                    // 科目编码
                    // 将对应的贷方科目替换为相应value
                    row.createCell(8).setCellValue(creditMinorNameJPInfo.get(info.getCreditMinorNameJP()));

                    // 11/16 生成对应凭证号
                    // 凭证类别编码
                    row.createCell(2).setCellValue("06");
                    // 凭证号
                    row.createCell(3).setCellValue("5");

                    // 代扣代缴分录 合并分录 生成一条
                    if (payForOtherGroup == null) {
                        payForOtherGroup = indexRow;
                        indexRow++;
                        // 生成一条借方科目为221101的数据 金额和借方相同(税前金额)
                        Row rowTmpCredit = sheet.createRow(payForOtherGroup);

                        // 共同信息生成
                        createCommon(rowTmpCredit, info);
                        // 11/16 生成对应凭证号
                        // 凭证类别编码
                        rowTmpCredit.createCell(2).setCellValue("06");
                        // 凭证号
                        rowTmpCredit.createCell(3).setCellValue("5");
                        //摘要
                        rowTmpCredit.createCell(7).setCellValue("代扣代缴分录");
                        //科目编码
                        rowTmpCredit.createCell(8).setCellValue("221101");
                        //原币贷方金额
                        Cell cell10 = rowTmpCredit.createCell(10);
                        cell10.setCellStyle(cellStyle);
                        cell10.setCellValue(info.getCreditCountWithTax());
                        //本币贷方金额
                        if (StringUtils.hasLength(info.getCreditCountWithTax())) {
                            BigDecimal creditCount = new BigDecimal(info.getCreditCountWithTax().replaceAll(",", ""));
                            BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                            Cell cell11 = rowTmpCredit.createCell(11);
                            cell11.setCellStyle(cellStyle);
                            cell11.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                        }
                    } else {
                        Row payForOtherGroupRow = sheet.getRow(payForOtherGroup);
                        //原币贷方金额
                        Cell cell10 = payForOtherGroupRow.getCell(10);
                        BigDecimal cell10Value = new BigDecimal(cell10.getStringCellValue().replaceAll(",", ""));
                        BigDecimal current10Value = new BigDecimal(info.getCreditCountWithTax().replaceAll(",", ""));
                        cell10.setCellValue(df.format(cell10Value.add(current10Value).setScale(2, RoundingMode.HALF_UP)));
                        //本币贷方金额
                        if (StringUtils.hasLength(info.getCreditCountWithTax())) {
                            BigDecimal creditCount = cell10Value.add(current10Value).setScale(2, RoundingMode.HALF_UP);
                            BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                            Cell cell11 = payForOtherGroupRow.getCell(11);
                            cell11.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                        }
                    }
                }

                // 辅助科目
                createMinorCode(info, row, "1");
            }

            // 如果消费税科目-贷方
            if (StringUtils.hasLength(info.getCreditConsumptionTax())) {
                Row row = sheet.createRow(indexRow);
                indexRow++;

                // 共同信息生成
                createCommon(row, info);
                row.createCell(7).setCellValue("消费税");
                //科目编码
                row.createCell(8).setCellValue(getTaxCodeByTaxInfo(info, "2"));
                //原币贷方金额
                Cell cell18 = row.createCell(18);
                cell18.setCellStyle(cellStyle);
                cell18.setCellValue(info.getCreditConsumptionTax());
                //本币贷方金额
                if (StringUtils.hasLength(info.getCreditConsumptionTax())) {
                    BigDecimal creditCount = new BigDecimal(info.getCreditConsumptionTax().replaceAll(",", ""));
                    BigDecimal exchange = new BigDecimal(info.getExchangeRate());

                    Cell cell19 = row.createCell(19);
                    cell19.setCellStyle(cellStyle);
                    cell19.setCellValue(df.format(creditCount.multiply(exchange).setScale(2, RoundingMode.HALF_UP)));
                }

                createCodeIndex(info, row);
            }
        }

        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

        response.setCharacterEncoding("utf-8");
        response.setHeader("content-Type", "application/vnd.ms-excel");
        workbook.write(response.getOutputStream());
    }

    // 共同生成逻辑
    private void createCommon(Row row, FinancialExcelInfo info) {
        row.createCell(1).setCellValue("100112-0002");
        //制单人编码
        row.createCell(5).setCellValue("shiyu");
        //制单日期
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        Calendar calendar = Calendar.getInstance();
        calendar.add(Calendar.MONTH, 0);
        calendar.set(Calendar.DAY_OF_MONTH, 0);
        String lastDay = format.format(calendar.getTime());
        row.createCell(6).setCellValue(lastDay);
        //币种
        row.createCell(9).setCellValue("日元");
        //业务单元编码
        row.createCell(14).setCellValue("100112");

        row.createCell(25).setCellValue(lastDay);
        //组织本币汇率
        row.createCell(37).setCellValue(info.getExchangeRate());

        // 是否为工资列
        if (StringUtils.hasLength(info.getIsSalary()) && info.getIsSalary().equals("1")) {
            row.createCell(49).setCellValue(info.getIsSalary());
        }

        // 生成日本凭证号 方便核对
        row.createCell(50).setCellValue(info.getIndex());
    }

    // 11/7 设计修改 根据不同的税区分 获取对应消费税科目
    private String getTaxCodeByTaxInfo(FinancialExcelInfo info, String flg) {
        // 如果消费税科目-借方
        if (flg.equals("1") && StringUtils.hasLength(info.getDebitMinorTaxInfo())) {
            // 1.課対仕入8%（軽）只要税区分列内容包含8% 科目:2221120101
            if (info.getDebitMinorTaxInfo().contains("8%")) {
                return "2221120101";
            }

            // 2.課対仕入10% 完全匹配 科目:2221120102
            if (info.getDebitMinorTaxInfo().equals("課対仕入10%")) {
                return "2221120102";
            }

            // 3.課税売上10% 完全匹配 科目:2221120202
            if (info.getDebitMinorTaxInfo().equals("課税売上10%")) {
                return "2221120202";
            }

        }

        // 如果消费税科目-贷方
        if (flg.equals("2") && StringUtils.hasLength(info.getCreditMinorTaxInfo())) {
            // 1.課対仕入8%（軽）只要税区分列内容包含8% 科目:2221120101
            if (info.getCreditMinorTaxInfo().contains("8%")) {
                return "2221120101";
            }

            // 2.課対仕入10% 完全匹配 科目:2221120102
            if (info.getCreditMinorTaxInfo().equals("課対仕入10%")) {
                return "2221120102";
            }

            // 3.課税売上10% 完全匹配 科目:2221120202
            if (info.getCreditMinorTaxInfo().equals("課税売上10%")) {
                return "2221120202";
            }
        }

        return "2221120102";
    }

    // 生成辅助科目
    private void createMinorCode(FinancialExcelInfo info, Row row, String flg) {
        String code = row.getCell(8).getStringCellValue();
        // 需要生成辅助科目
        if (codeListForMinorCode.containsKey(code)) {
            // 部门
            String minorDepartmentName = info.getDebitMinorDepartmentName();
            // 日方提供辅助科目
            String minorNameJP = info.getDebitMinorNameJP();

            // 科目为贷方情况
            if (flg.equals("1")) {
                minorDepartmentName = info.getCreditMinorDepartmentName();
                minorNameJP = info.getCreditMinorNameJP();
            }

            switch (codeListForMinorCode.get(code)) {
                // 科目code 2202
                case "2202":
                    // 辅助科目 部门情况
                    if (StringUtils.hasLength(minorDepartmentName)) {
                        //辅助核算1 部门情况
                        row.createCell(40).setCellValue(minorDepartmentName.concat(":部门"));
                        //辅助核算2 默认
                        row.createCell(41).setCellValue("9999:合同档案(自定义档案)");
                        //辅助核算3 供应商
                        if (minorNameJP.contains("上海新致") || minorNameJP.contains("大连新致")) {
                            row.createCell(42).setCellValue("1001:供应商档案");
                        } else {
                            row.createCell(42).setCellValue("S_".concat(minorNameJP).concat(":供应商档案"));
                        }
                        //辅助核算4 部门情况
                        row.createCell(43).setCellValue(minorDepartmentName.concat(":成本中心(自定义档案)"));
                    }
                    break;
                // 科目code 100502 银行账户
                case "100502":
                    //辅助核算1 银行账户
                    if (StringUtils.hasLength(minorNameJP)) {
                        row.createCell(40).setCellValue(minorCodeForBank.get(minorNameJP));
                    }
                    break;
                // 科目code 660301 银行账户
                case "660301":
                    //辅助核算1 银行账户
                    if (minorNameJP.contains("東日本")) {
                        row.createCell(40).setCellValue("325798:银行账户");
                    }
                    if (minorNameJP.contains("みずほ")) {
                        row.createCell(40).setCellValue("2226709:银行账户");
                    }
                    break;
                // 科目code 112203
                case "112203":
                    // 科目code 600103
                case "600103":
                    // 辅助科目 部门情况
                    if (StringUtils.hasLength(minorDepartmentName)) {
                        //辅助核算1 部门情况
                        row.createCell(40).setCellValue(minorDepartmentName.concat(":部门"));
                        //辅助核算2 默认
                        row.createCell(41).setCellValue("9999:合同档案(自定义档案)");
                        //辅助核算3 默认
                        row.createCell(42).setCellValue("88888:项目档案(自定义档案)");
                        //辅助核算4 客户档案
                        row.createCell(43).setCellValue("C_".concat(minorNameJP).concat(":客户档案"));
                        //辅助核算5 部门情况
                        row.createCell(44).setCellValue(minorDepartmentName.concat(":成本中心(自定义档案)"));
                    }
                    break;
                // 科目code 112213
                case "112213":
                    // 辅助科目 部门情况
                    if (StringUtils.hasLength(minorDepartmentName)) {
                        //辅助核算1 部门情况
                        row.createCell(40).setCellValue(minorDepartmentName.concat(":部门"));
                        //辅助核算2 默认
                        row.createCell(41).setCellValue("9999:合同档案(自定义档案)");
                        //辅助核算3 默认
                        row.createCell(42).setCellValue("88888:项目档案(自定义档案)");
                        //辅助核算4 部门情况
                        row.createCell(43).setCellValue(minorDepartmentName.concat(":成本中心(自定义档案)"));
                        //辅助核算5 客户档案
                        row.createCell(44).setCellValue("C_".concat(minorNameJP).concat(":客户档案"));
                    }
                    break;
                // 科目code 500130
                case "500130":
                    // 科目code 500131
                case "500131":
                    // 辅助科目 部门情况
                    if (StringUtils.hasLength(minorDepartmentName)) {
                        //辅助核算1 部门情况
                        row.createCell(40).setCellValue(minorDepartmentName.concat(":部门"));
                        //辅助核算2 默认
                        row.createCell(41).setCellValue("88888:项目档案(自定义档案)");
                        //辅助核算3 客户档案
                        if (minorNameJP.contains("上海新致") || minorNameJP.contains("大连新致")) {
                            row.createCell(42).setCellValue("1001:供应商档案");
                        } else {
                            row.createCell(42).setCellValue("S_".concat(minorNameJP).concat(":供应商档案"));
                        }
                        //辅助核算4 部门情况
                        row.createCell(43).setCellValue(minorDepartmentName.concat(":成本中心(自定义档案)"));
                    }
                    break;
                // 科目code 122103
                case "122103":
                    // 辅助科目 客户档案
                    if (StringUtils.hasLength(minorNameJP)) {
                        //辅助核算1 客户档案
                        row.createCell(40).setCellValue("C_".concat(minorNameJP).concat(":客户档案"));
                    }
                    break;
                // 科目code 224101
                case "224101":
                    // 辅助科目 供应商档案
                    if (StringUtils.hasLength(minorNameJP)) {
                        //辅助核算1 供应商档案
                        if (minorNameJP.contains("上海新致") || minorNameJP.contains("大连新致")) {
                            row.createCell(40).setCellValue("1001:供应商档案");
                        } else {
                            row.createCell(40).setCellValue("S_".concat(minorNameJP).concat(":供应商档案"));
                        }
                    }
                    break;
                // 科目code 224102
                case "224102":
                    // 辅助科目 工会费:供应商档案
                    row.createCell(40).setCellValue("S_工会费:供应商档案");
                    break;
                // 科目code 224104
                case "224104":
                    // 辅助科目 社保:供应商档案
                    row.createCell(40).setCellValue("S_社保:供应商档案");
                    break;
                // 科目code 500101
                case "500101":
                    // 科目code 500103
                case "500103":
                    // 辅助科目 部门情况
                    if (StringUtils.hasLength(minorDepartmentName)) {
                        //辅助核算1 部门情况
                        row.createCell(40).setCellValue(minorDepartmentName.concat(":部门"));
                        //辅助核算2 默认
                        row.createCell(41).setCellValue("88888:项目档案(自定义档案)");
                        //辅助核算3 部门情况
                        row.createCell(42).setCellValue(minorDepartmentName.concat(":成本中心(自定义档案)"));
                    }
                    break;
                // 默认按照部门生成
                default:
                    // 辅助科目 部门情况
                    if (StringUtils.hasLength(minorDepartmentName)) {
                        //辅助核算1 部门情况
                        row.createCell(40).setCellValue(minorDepartmentName.concat(":部门"));
                        //辅助核算2 部门情况
                        row.createCell(41).setCellValue(minorDepartmentName.concat(":成本中心(自定义档案)"));
                    }
                    break;
            }
        }
    }

    // 生成凭证号
    private void createCodeIndex(FinancialExcelInfo info, Row row) {
        // 判断非工资情况
        if (!StringUtils.hasLength(info.getIsSalary())) {
            //如果和当前伝票No.相同则为一组 使用统一的凭证号
            if (info.getIndex().equals(currentIndex)) {
                rows.add(row);
            } else {
                // 将之前伝票No.组按照优先级最高的规则 进行批量修改
                if (StringUtils.hasLength(currentIndex)) {
                    for (Row detail : rows) {
                        // 凭证类别编码
                        detail.createCell(2).setCellValue(typeCode);
                        // 凭证号
                        detail.createCell(3).setCellValue(code);
                    }
                }

                // 新的伝票No.组 生成
                currentIndex = info.getIndex();
                rows.clear();
                // 当前凭证号生成规则优先级
                weight = 10;
                // 凭证类型编码
                typeCode = "";
                // 凭证号
                code = "";
                rows.add(row);
            }

            // 调用具体生成逻辑
            createCodeIndexLogic(row, info);
        }
    }

    // 生成凭证号 具体逻辑
    private void createCodeIndexLogic(Row row, FinancialExcelInfo info) {
        // 获取科目列
        String resultCode = row.getCell(8).getStringCellValue();
        // 获取借方金额列
        String debitTmp = "";
        if (row.getCell(10) != null) {
            debitTmp = row.getCell(10).getStringCellValue();
        }
        // 获取贷方金额列
        String creditTmp = "";
        if (row.getCell(18) != null) {
            creditTmp = row.getCell(18).getStringCellValue();
        }

        //1.（客户回款）银收1：（112213）应收账款，在贷方。
        if (StringUtils.hasLength(creditTmp) && resultCode.equals("112213")) {
            // 判断优先级
            if (1 < weight) {
                weight = 1;
                // 凭证类型编码
                typeCode = "02";
                // 凭证号
                code = "1";
            }
        }

        //2.（确认收入）转账1：（600103）主营业务收入-服务收入-日元，在贷方
        else if (StringUtils.hasLength(creditTmp) && resultCode.equals("600103")) {
            // 判断优先级
            if (2 < weight) {
                weight = 2;
                // 凭证类型编码
                typeCode = "06";
                // 凭证号
                code = "1";
            }
        }

        //3.（计提内、外包费）转账2：（2202）应付账款，在贷方
        else if (StringUtils.hasLength(creditTmp) && resultCode.equals("2202")) {
            // 判断优先级
            if (3 < weight) {
                weight = 3;
                // 凭证类型编码
                typeCode = "06";
                // 凭证号
                code = "2";
            }
        }

        //4.（支付内、外包费）银付1：（2202）应付账款，在借方。
        else if (StringUtils.hasLength(debitTmp) && resultCode.equals("2202")) {
            // 判断优先级
            if (4 < weight) {
                weight = 4;
                // 凭证类型编码
                typeCode = "03";
                // 凭证号
                code = "1";
            }
        }

        //5.（现金支付报销）现付1：（100103）库存现金，在贷方分录。
        else if (StringUtils.hasLength(creditTmp) && resultCode.equals("100103")) {
            // 判断优先级
            if (5 < weight) {
                weight = 5;
                // 凭证类型编码
                typeCode = "05";
                // 凭证号
                code = "1";
            }
        }

        //6.（银行支付报销）银付2：（100502）银行存款，在贷方，按照日本凭证区分编号
        else if (StringUtils.hasLength(creditTmp) && resultCode.equals("100502")) {
            // 判断优先级
            if (6 < weight) {
                weight = 6;
                // 凭证类型编码
                typeCode = "03";
                // 凭证号
                code = "2";
            }
        }

        //7.现收1，（100103）库存现金在借方，正常排号
        else if (StringUtils.hasLength(debitTmp) && resultCode.equals("100103")) {
            // 判断优先级
            if (7 < weight) {
                weight = 7;
                // 凭证类型编码
                typeCode = "04";
                // 凭证号
                code = "1";
            }
        }

        //8.银收2，（100502）银行存款，在借方
        else if (StringUtils.hasLength(debitTmp) && resultCode.equals("100502")) {
            // 判断优先级
            if (8 < weight) {
                weight = 8;
                // 凭证类型编码
                typeCode = "02";
                // 凭证号
                code = "2";
            }
        }

        //9. 其他
        else {
            // 判断优先级
            if (9 < weight) {
                weight = 9;
                // 凭证类型编码
                typeCode = "06";
                // 凭证号
                code = "3";
            }
        }
    }

    // 重置初始化数据
    private void reset() {
        // 当前伝票No.
        currentIndex = "";
        // 当前凭证号生成规则优先级
        weight = 10;
        // 凭证类型编码
        typeCode = "";
        // 凭证号
        code = "";
        // 当前相同伝票No.的数据组
        rows = new ArrayList<Row>();

        // 分录合并
        // 计提工资分录
        salaryGroup = null;
        // 支付工资分录
        paySalaryGroup = null;
        // 代扣代缴
        payForOtherGroup = null;
        // 计提保险
        insuranceGroup = null;
    }
}
