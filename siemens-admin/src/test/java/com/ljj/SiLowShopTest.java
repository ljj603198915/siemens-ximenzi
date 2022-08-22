package com.ljj;

import com.lheia.common.exception.base.BaseException;
import com.lheia.common.utils.StringUtils;
import com.lheia.common.utils.bean.DtoEntityUtil;
import com.lheia.common.utils.poi.ExcelUtilLh;
import com.lheia.icm.domain.db.*;
import com.lheia.icm.domain.dto.importIc.GscImportDto;
import com.lheia.icm.domain.dto.importIc.IcgImportDto;
import com.lheia.icm.domain.dto.importIc.InvestClassImportDto;
import com.lheia.icm.mapper.*;
import com.lheia.icm.service.IGreenStandardsClassService;
import com.lheia.icm.service.IInvestClassGroupService;
import com.ljj.common.exception.base.BaseException;
import com.ljj.common.utils.StringUtils;
import com.ljj.common.utils.poi.ExcelUtilLh;
import com.ljj.domain.Lowshop;
import com.ljj.mapper.LowshopMapper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@SpringBootTest(classes = RuoYiApplication.class)
public class SiLowShopTest {

    @Autowired
    private LowshopMapper lowshopMapper;


    @Test
//    public void importInvestClassGroupTest() {
//        InvestClassGroup investClassGroup = investClassGroupMapper.selectById(1L);
//
//        String pathName = "D:\\项目\\南海\\小五类维护\\小五类相关文档-20220726\\李俊杰测试\\环境效益测算模型更新-2022-06-07(4)(1).xlsx";
//        File file = new File(pathName);
//
//        Workbook wb = null;
//        IcgImportDto icgImportDto = null;
//        List<GreenStandards> gsList = greenStandardsMapper.selectGreenStandardsList(null);
//        try {
//            wb = ExcelUtilLh.getWorkbok(file);
//            Sheet sheet = wb.getSheetAt(0);
//            icgImportDto = investClassGroupService.readFileInvestClass(sheet, investClassGroup.getCode(), gsList);
//            System.out.println("数据" + icgImportDto);
//        } catch (Exception e) {
//            System.out.println("文件读取失败" + e.getMessage());
//        }
//        addR(investClassGroup, icgImportDto);
//    }

    public List<Lowshop> readExcel(Sheet sheet) {
        if (null == sheet) return null;
        ArrayList<Lowshop> resList = new ArrayList<>();
        try {
            int lastRowNum = sheet.getLastRowNum(); //获取内容占用的最后一行（以0开始）

            //处理导入内容
            for (int j = 0; j <= lastRowNum; j++) {
                //获取每行
                Row row = sheet.getRow(j);
                if (null == row) continue;
                int headNum = 0;//头
                if (j > headNum) {
                    String code = "";
                    String codeShow = "";
                    int level = 0;
                    String benefitCode = null;
                    String longitude = "";
                    String latitude = "";
                    String shopName = ExcelUtilLh.getStringCellValue(row.getCell(0));
                    String province = ExcelUtilLh.getStringCellValue(row.getCell(1));
                    String city = ExcelUtilLh.getStringCellValue(row.getCell(2));
                    String district = ExcelUtilLh.getStringCellValue(row.getCell(3));
                    String adress = ExcelUtilLh.getStringCellValue(row.getCell(4));
                    String telephone = ExcelUtilLh.getStringCellValue(row.getCell(5));
                    String longLat = ExcelUtilLh.getStringCellValue(row.getCell(6));//关键字
                    String shopType = ExcelUtilLh.getStringCellValue(row.getCell(7));//绿色投向说明
                    String onSale = ExcelUtilLh.getStringCellValue(row.getCell(9));//上传文件说明
                    String resTypeId = ExcelUtilLh.getStringCellValue(row.getCell(9));//上传文件说明
                    String displayOrder = ExcelUtilLh.getStringCellValue(row.getCell(9));//上传文件说明

                    if (StringUtils.isNotEmpty(longLat)) {
                        longitude = longLat.split(",")[0];
                        latitude = longLat.split(",")[1];
                    }

                    Lowshop lowshop = new Lowshop();
                    lowshop.setShopName(shopName);
                    lowshop.setProvince(province);
                    lowshop.setCity(city);
                    lowshop.setDistrict(district);
                    lowshop.setAdress(adress);
                    lowshop.setTelephone(telephone);
                    lowshop.setLongitude(longitude);
                    lowshop.setLatitude(latitude);
                    resList.add(lowshop);
                }
            }

        } catch (Exception e) {
            System.out.println("读取附近店铺数据异常" + e.getMessage());
            throw new BaseException("读取附近店铺数据异常" + e.getMessage());
        }

        return resList;
    }

    public static Integer mappingShopType(String shopType){

    }
}
