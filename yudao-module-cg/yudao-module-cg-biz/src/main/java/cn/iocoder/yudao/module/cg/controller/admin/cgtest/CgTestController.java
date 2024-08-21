package cn.iocoder.yudao.module.cg.controller.admin.cgtest;

import cn.hutool.core.collection.CollUtil;
import cn.iocoder.yudao.framework.common.pojo.CommonResult;
import cn.iocoder.yudao.framework.common.pojo.PageResult;
import cn.iocoder.yudao.framework.common.util.object.BeanUtils;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgPagePeqVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgRespVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgSaveReqVO;
import cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo.CgUpdateReqVO;
import cn.iocoder.yudao.module.cg.convert.MinioUtil;
import cn.iocoder.yudao.module.cg.dal.dataobject.cgtest.CgTestDO;
import cn.iocoder.yudao.module.cg.framework.minio.config.MinioProperties;
import cn.iocoder.yudao.module.cg.service.cgtest.CgTestService;
import cn.iocoder.yudao.module.system.api.user.AdminUserApi;
import cn.iocoder.yudao.module.system.api.user.dto.AdminUserRespDTO;
import com.fasterxml.jackson.databind.ObjectMapper;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.annotation.Resource;
import jakarta.validation.Valid;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static cn.iocoder.yudao.framework.common.pojo.CommonResult.success;

@Tag(name = "管理后台 - Test")
@RestController
@RequestMapping("/cg/test")
@Validated
@Slf4j
public class CgTestController {

    @Resource
    private MinioProperties minioProperties;

    @Autowired
    private MinioUtil minioUtil;

    @Resource
    private CgTestService cgTestService;

    @Resource
    private AdminUserApi adminUserApi;
    @PostMapping("/create")
    @Operation(summary = "新增用户")
    public CommonResult<Long> createUser(@Valid @RequestBody CgSaveReqVO reqVO) {
        Long id = cgTestService.createUser(reqVO);
        return success(id);
    }

    @PutMapping("/update")
    @Operation(summary = "修改用户")
    public CommonResult<Boolean> updateUser(@Valid @RequestBody CgUpdateReqVO reqVO) {
        cgTestService.updateUser(reqVO);
        return success(true);
    }
    @DeleteMapping("/delete")
    @Operation(summary = "删除用户")
    @Parameter(name = "id", description = "编号", required = true, example = "1024")
//    @PreAuthorize("@ss.hasPermission('system:user:delete')")
    public CommonResult<Boolean> deleteUser(@RequestParam("id") Long id) {
        cgTestService.deleteUser(id);
        return success(true);
    }
    @DeleteMapping("/delete-all")
    @Operation(summary = "批量删除用户")
    @Parameter(name = "ids", description = "编号列表,逗号分割", required = true, example = "10,11,12")
//    @PreAuthorize("@ss.hasPermission('system:user:delete')")
    public CommonResult<Boolean> deleteUsers(@RequestParam("ids") String ids) {
        String[] items = ids.split(",");
        List<Long> list = new ArrayList<>();
        for (String item : items) {
            list.add(Long.parseLong(item));
        }
        cgTestService.deleteUsers(list);
        return success(true);
    }

    @GetMapping("/get")
    @Operation(summary = "获得用户详情")
    @Parameter(name = "id", description = "编号", required = true, example = "1024")
    public CommonResult<CgRespVO> getUser(@RequestParam("id") Long id) {
        CgTestDO user = cgTestService.getUser(id);
        if (user == null) {
            return success(null);
        }
        LocalDateTime createTime1 = user.getCreateTime();
        CgRespVO cgRespVO = BeanUtils.toBean(user, CgRespVO.class);
        LocalDateTime createTime = cgRespVO.getCreateTime();
        log.error("createTime1=="+createTime1);
        log.error("createTime=="+createTime);
        long creatorId = Long.parseLong(user.getCreator());
        long updaterId = Long.parseLong(user.getUpdater());
        AdminUserRespDTO creator = adminUserApi.getUser(creatorId);
        AdminUserRespDTO updater = adminUserApi.getUser(updaterId);
        cgRespVO.setCreatorName(creator.getNickname());
        cgRespVO.setUpdaterName(updater.getNickname());

        return success(cgRespVO);
    }

    @GetMapping("/get-all")
    @Operation(summary = "获得用户分页列表")
    public CommonResult<PageResult<CgRespVO>> getUsers(@Valid CgPagePeqVO pagePeqVO) {
        // 获得用户分页列表
        PageResult<CgTestDO> pageResult = cgTestService.getUserPage(pagePeqVO);
        if (CollUtil.isEmpty(pageResult.getList())) {
            return success(new PageResult<>(pageResult.getTotal()));
        }
        List<CgTestDO> list = pageResult.getList();
        List<CgRespVO> cgRespVOList = new ArrayList<>();
        for (CgTestDO cgTestDO:list){
            CgRespVO cgRespVO = BeanUtils.toBean(cgTestDO, CgRespVO.class);

            long creatorId = Long.parseLong(cgTestDO.getCreator());
            long updaterId = Long.parseLong(cgTestDO.getUpdater());
            AdminUserRespDTO creator = adminUserApi.getUser(creatorId);
            AdminUserRespDTO updater = adminUserApi.getUser(updaterId);
            cgRespVO.setCreatorName(creator.getNickname());
            cgRespVO.setUpdaterName(updater.getNickname());

            cgRespVOList.add(cgRespVO);
        }

        return success(new PageResult<>(cgRespVOList,pageResult.getTotal()));
    }
    @GetMapping("/get-demo")
    @Operation(summary = "获取 test 信息")
    public CommonResult<Boolean> get() {

        System.out.println("minioProperties.getBucketName() = " + minioProperties.getBucketName());
        System.out.println("minioProperties.getEndpoint() = " + minioProperties.getEndpoint());
        System.out.println("minioProperties.getAccessKey() = " + minioProperties.getAccessKey());
        System.out.println("minioProperties.getSecretKey() = " + minioProperties.getSecretKey());
        System.out.println("minioProperties.getReportTemp() = " + minioProperties.getReportTemp());
        String modelCodeFmtdPath = "/cg/test/通用封面（套打）-食品_20201016154838.xlsx";
        String localfmtdPath = "/Users/chengang/Documents/cg.xlsx";
        minioUtil.downloadFile(minioProperties.getBucketName(),modelCodeFmtdPath,localfmtdPath);
        System.out.println("localfmtdPath=="+localfmtdPath);

//        // 读取Excel文件
//        InputStream fis = null;
//        try {
//            fis = new FileInputStream("/Users/chengang/Documents/ss.xlsx");
//            Workbook workbook = WorkbookFactory.create(fis);
//            Sheet sheet = workbook.getSheetAt(0);
//
//            // 定义关键字和图片路径的映射关系
//            Map<String, String> imageMap = new HashMap<>();
//            imageMap.put("&{主检}", "/Users/chengang/Documents/stestby.png");
//            imageMap.put("&{二审}", "/Users/chengang/Documents/sapprovedby.png");
//            imageMap.put("&{终审}", "/Users/chengang/Documents/sreleasedby.png");
//            imageMap.put("(检验检测专用章)", "/Users/chengang/Documents/stamp.png");
//
//            // 插入图片
//            insertImages(workbook, sheet, imageMap);
//
//            // 写回到新的Excel文件
//            FileOutputStream fos = new FileOutputStream("/Users/chengang/Documents/updated_61.xlsx");
//            workbook.write(fos);
//            fos.close();
//            workbook.close();
//            fis.close();
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        }

//        String apiUrl = "http://apitest.cnqos.com:8001/api/PREORDERS/GetGenerateReportInforData";
//        String preordNo = "NZJ(2023)SP01-34095";
//        String tempCode = "165";
//        try {
//            // 创建HttpClient
//            HttpClient client = HttpClient.newHttpClient();
//
//            // 构建请求URL
//            String fullUrl = apiUrl + "?preordNo=" + preordNo + "&tempCode=" + tempCode;
//
//            // 创建POST请求
//            HttpRequest request = HttpRequest.newBuilder()
//                    .uri(URI.create(fullUrl))
//                    .GET()
//                    .header("Content-Type", "application/json")
//                    .build();
//
//            // 发送请求并获取响应
//            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString(StandardCharsets.UTF_8));
//
//            // 检查响应状态码
//            if (response.statusCode() == 200) {
//                String responseBody = response.body();
//                Map<String, String> responseMap = parseResponse(responseBody);
//                System.out.println("Parsed response:");
//                responseMap.forEach((key, value) -> System.out.println(key + ": " + value));
//                String excelPath = "/Users/chengang/Downloads/食品中心委托南京市质检院首页模板_20240205142318.xls"; // 请将此路径替换为您的实际文件路径
//                fillExcel(excelPath,responseMap);
//            } else {
//                System.out.println("Request failed. Status code: " + response.statusCode());
//            }
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
        return success(true);
    }
    private  Map<String, String> parseResponse(String jsonResponse) throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        Map<String, Object> jsonMap = mapper.readValue(jsonResponse, Map.class);

        if (jsonMap.containsKey("response") && jsonMap.get("response") instanceof Map) {
            @SuppressWarnings("unchecked")
            Map<String, String> responseMap = (Map<String, String>) jsonMap.get("response");
            return responseMap;
        } else {
            throw new Exception("Response does not contain expected 'response' object");
        }
    }
    public void fillExcel(String excelPath, Map<String, String> dataMap) {
        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new HSSFWorkbook(fis)) {
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            for (Map.Entry<String, String> entry : dataMap.entrySet()) {
                                String placeholder = entry.getKey();
                                String replacement = entry.getValue();
                                cellValue = cellValue.replace(placeholder, replacement);
                            }
                            cell.setCellValue(cellValue);
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

    /**
     * 将图片插入到Excel中的指定位置，保持图片比例不变
     *
     * @param workbook  Excel工作簿
     * @param sheet     工作表
     * @param imageMap  关键字和图片路径的映射关系
     * @throws IOException 读取图片时的异常
     */
    public void insertImages(Workbook workbook, Sheet sheet, Map<String, String> imageMap) throws IOException {
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
                        System.out.println("mergedRegion.getFirstColumn() = " + mergedRegion.getFirstColumn());
                        System.out.println("mergedRegion.getLastColumn() = " + mergedRegion.getLastColumn());
                        System.out.println("mergedRegion.getFirstRow() = " + mergedRegion.getFirstRow());
                        System.out.println("mergedRegion.getLastRow() = " + mergedRegion.getLastRow());
                        // 计算合并单元格的总宽度
                        float cellWidth = 0;
                        for (int col = mergedRegion.getFirstColumn(); col <= mergedRegion.getLastColumn(); col++) {
                            cellWidth += sheet.getColumnWidthInPixels(col);
                        }

                        // 计算合并单元格的总高度
                        float cellHeight = 0;
                        for (int r = mergedRegion.getFirstRow(); r <= mergedRegion.getLastRow(); r++) {
                            cellHeight += sheet.getRow(r).getHeightInPoints() / 72 * 96; // 转换为像素
                        }

                        InputStream is = new FileInputStream(imagePath);
                        byte[] bytes = IOUtils.toByteArray(is);
                        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                        is.close();
                        // 获取图片的原始宽高
                        BufferedImage image = ImageIO.read(new FileInputStream(imagePath));
                        int pictureWidth = image.getWidth();
                        int pictureHeight = image.getHeight();

                        // 计算缩放比例，保持图片的宽高比一致
                        double scale = Math.min(cellWidth / pictureWidth, cellHeight / pictureHeight);
                        System.out.println("scale = " + scale);
                        // 计算缩放后的图片宽高
                        int newWidth = (int) (pictureWidth * scale);
                        int newHeight = (int) (pictureHeight * scale);

                        // 计算偏移量，使图片居中
                        int dx1 = (int) ((cellWidth - newWidth) / 2);
                        int dy1 = (int) ((cellHeight - newHeight) / 2);
//                        int dx2 = dx1 + newWidth;
//                        int dy2 = dy1 + newHeight;

                        // 设置锚点并插入图片
                        ClientAnchor anchor = helper.createClientAnchor();
                        anchor.setCol1(mergedRegion.getFirstColumn());
                        anchor.setRow1(mergedRegion.getFirstRow());
                        anchor.setDx1(dx1 * Units.EMU_PER_PIXEL);
                        anchor.setDy1(dy1 * Units.EMU_PER_PIXEL);
//                        anchor.setDx2(dx2 * Units.EMU_PER_PIXEL);
//                        anchor.setDy2(dy2 * Units.EMU_PER_PIXEL);
//                        anchor.setCol2(mergedRegion.getLastColumn());
//                        anchor.setRow2(mergedRegion.getLastRow());

                        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

                        Picture picture = drawing.createPicture(anchor,pictureIdx);
                        picture.resize(scale);
                        //固定好了dx1，dy1，dx2，dy2，然后就可以调整宽高了，不需要使用resize方法缩放
//                        //设置单元格居中
//                        CellStyle style = workbook.createCellStyle();
//                        style.setAlignment(HorizontalAlignment.CENTER);
//                        style.setVerticalAlignment(VerticalAlignment.CENTER);
//                        cell.setCellStyle(style);
                    }
                }
            }
        }
    }

    /**
     * 获取合并单元格的范围
     *
     * @param sheet  工作表
     * @param row    单元格行号
     * @param column 单元格列号
     * @return 合并单元格的范围
     */
    private CellRangeAddress getMergedRegion(Sheet sheet, int row, int column) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(row, column)) {
                return merged;
            }
        }
        return new CellRangeAddress(row, row, column, column);
    }
}
