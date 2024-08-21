package cn.iocoder.yudao.module.cg.controller.admin.excel;

import cn.iocoder.yudao.framework.common.pojo.CommonResult;
import cn.iocoder.yudao.module.cg.controller.admin.excel.vo.*;
import cn.iocoder.yudao.module.cg.convert.ExcelMthd;
import cn.iocoder.yudao.module.cg.convert.FtpUtil;
import cn.iocoder.yudao.module.cg.convert.MinioUtil;
import cn.iocoder.yudao.module.cg.convert.UtilTools;
import cn.iocoder.yudao.module.cg.framework.minio.config.MinioProperties;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.annotation.Resource;
import jakarta.validation.Valid;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.net.ftp.FTPClient;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.*;

import static cn.iocoder.yudao.framework.common.pojo.CommonResult.success;


@Tag(name = "管理后台 - Excel")
@RestController
@RequestMapping("/cg/excel")
@Validated
@Slf4j
public class ExcelController {

    @Resource
    private MinioProperties minioProperties;
    @Autowired
    private FtpUtil ftpUtil;

    @Autowired
    private MinioUtil minioUtil;

    @Autowired
    private ExcelMthd excelMthd;

    @Autowired
    private UtilTools utils;

    @Value("${yudao.qrcode-url}")
    private String qrcodeUrl;

    @PostMapping("/create")
    @Operation(summary = "创建3个excel")
    public CommonResult<CreateRespVO> createOneThreeExcelAndMerge(@Valid @RequestBody CreateExcel createExcelVO) {

        String folderno = createExcelVO.getFOLDERNO();
        String preordno = createExcelVO.getPREORDNO();
        String reportgentype = createExcelVO.getREPORTGENTYPE();
        String inspectiondate = createExcelVO.getINSPECTIONDATE();
        String qualificationsfordepartment = createExcelVO.getQUALIFICATIONSFORDEPARTMENT();
        String modelCodeFmtdPath = createExcelVO.getMODEL_CODE_FMTD_PATH();
        String modelCodeFmPath = createExcelVO.getMODEL_CODE_FM_PATH();
        String modelCodeSyPath = createExcelVO.getMODEL_CODE_SY_PATH();
        //reportgentype:自动的时候用到这个字段，需要填充
        String modelCodeFyPath = createExcelVO.getMODEL_CODE_FY_PATH();
        String modelCode2Path = createExcelVO.getMODEL_CODE2_PATH();
        String modelCode3Path = createExcelVO.getMODEL_CODE3_PATH();
        //reportgentype:合成的时候用到这个字段，直接下载附页
        String excelFyFtp = createExcelVO.getEXCEL_FY_FTP();
        String attachedPdfName = createExcelVO.getAttached_PDF_Name();
        Map<String, String> fmtdDataList = createExcelVO.getFMTDDataList();
        Map<String, String> fmDataList = createExcelVO.getFMDataList();
        Map<String, String> syDataList = createExcelVO.getSYDataList();
        FYDataModel fyDataModel = createExcelVO.getFYDataModel();

        System.out.println("folderno = " + folderno);
        System.out.println("fmtdDataList = " + fmtdDataList.toString());

        //创建本地目录
        String localPath = utils.localpath;

        utils.createLocalDirectory(localPath);

        String localfmtdPath = null;
        String localfmPath = null;
        String localsyPath = null;
        String localfyPath = null;
        String excelfyPath = null;
        String local2Path = null;
        String local3Path = null;

        //下载ftp文件
        //FTPClient ftpClient = ftpUtil.initializeFTPClient();
        //封页套打
        if (modelCodeFmtdPath != null && !modelCodeFmtdPath.isEmpty()) {

            //  /ReportTemplate/通用封面（套打）-食品_20201016154838.xls
//            String remotePath = modelCodeFmtdPath.substring(0, modelCodeFmtdPath.lastIndexOf("/"));
//            String fileName = modelCodeFmtdPath.substring(modelCodeFmtdPath.lastIndexOf("/") + 1);
//            System.out.println("remotePath = " + remotePath);
//            System.out.println("fileName = " + fileName);
            //创建下载模版的目录
            localfmtdPath = localPath + modelCodeFmtdPath;
            minioUtil.downloadFile(minioProperties.getBucketName(),modelCodeFmtdPath,localfmtdPath);
            System.out.println("localfmtdPath=="+localfmtdPath);
        }

        //封页
        if (modelCodeFmPath != null && !"".equals(modelCodeFmPath)) {

            localfmPath = localPath + modelCodeFmPath;
            minioUtil.downloadFile(minioProperties.getBucketName(),modelCodeFmPath,localfmPath);
            System.out.println("localfmPath=="+localfmPath);
        }
        //首页
        if (modelCodeSyPath != null && !"".equals(modelCodeSyPath)) {

            localsyPath = localPath + modelCodeSyPath;
            minioUtil.downloadFile(minioProperties.getBucketName(),modelCodeSyPath,localsyPath);
            System.out.println("localsyPath=="+localsyPath);
        }
        //附页
        if (modelCodeFyPath != null && !"".equals(modelCodeFyPath)) {

            localfyPath = localPath + modelCodeFyPath;
            minioUtil.downloadFile(minioProperties.getBucketName(),modelCodeFyPath,localfyPath);
            System.out.println("localfyPath=="+localfyPath);
        }
        //合成好的附页excel
        if (excelFyFtp != null && !"".equals(excelFyFtp)) {

            excelfyPath = localPath + excelFyFtp;
            minioUtil.downloadFile(minioProperties.getBucketName(),excelFyFtp,excelfyPath);
            System.out.println("excelfyPath=="+excelfyPath);
        }

        String remotePath = "/" + inspectiondate + "/" + folderno;
        CreateRespVO createRespVO = new CreateRespVO();
        //封页套打
        if (localfmtdPath != null && !localfmtdPath.isEmpty()) {
            int sheetIndex = 0;
            //填充数据
            excelMthd.reportStaticExcel(localfmtdPath, sheetIndex, fmtdDataList);
            //生成条形码、二维码
            String barcodePath = localPath + "/barcode/"+ preordno;
            utils.createLocalDirectory(barcodePath);
            excelMthd.generateBarcode(preordno, barcodePath+ "/barcode.png", 200, 40);
            excelMthd.generateQRCode(qrcodeUrl + preordno, barcodePath+ "/qrcode.png", 100, 100);
            //添加条形码
            String[] images = new String[1];
            images[0] = barcodePath + "/barcode.png";
            //初始化int二维数组
            int[][] imagesPoint = new int[images.length][4];
            imagesPoint[0][0] = 1;
            imagesPoint[0][1] = 24;
            imagesPoint[0][2] = 4;
            imagesPoint[0][3] = 25;
            double[] scaleFactors = new double[images.length];
            scaleFactors[0] = 0.7;
            excelMthd.insertImageIntoExcelCenter(localfmtdPath, 0, images, imagesPoint,scaleFactors);

            //添加二维码
            String[] qrImage = new String[1];
            qrImage[0] = barcodePath + "/qrcode.png";
            imagesPoint = new int[qrImage.length][4];
            imagesPoint[0][0] = 5;
            imagesPoint[0][1] = 24;
            imagesPoint[0][2] = 6;
            imagesPoint[0][3] = 26;
            scaleFactors = new double[qrImage.length];
            scaleFactors[0] = 1.0;
            excelMthd.insertImagesCenter(localfmtdPath, 0, qrImage, imagesPoint,scaleFactors);

//            //拷贝到文件服务器
//            try {
//                utils.copyFile(localfmtdPath, uploadPath,"fmtd.xlsx");
//            } catch (IOException e) {
//                throw new RuntimeException(e);
//            }
            String objectName = remotePath + "/fmtd.xlsx";
            minioUtil.uploadLocalFile(minioProperties.getBucketName(),localfmtdPath,objectName);

            createRespVO.setFmtdExcelPath(objectName);
//            // excel转pdf
//            utils.excel2pdf(preordno, "fmtd" ,uploadPath);
        }
        //封页
        if (localfmPath != null && !localfmPath.isEmpty()) {
            int sheetIndex = 0;
            //填充数据
            excelMthd.reportStaticExcel(localfmPath, sheetIndex, fmDataList);
            //生成条形码、二维码
            String barcodePath = localPath + "/barcode/"+ preordno;
            utils.createLocalDirectory(barcodePath);
            excelMthd.generateBarcode(preordno, barcodePath+ "/barcode.png", 200, 40);
            excelMthd.generateQRCode(qrcodeUrl + preordno, barcodePath+ "/qrcode.png", 100, 100);
            //添加图片
            //添加标
            String[] images = new String[1];
            images[0] = localPath + "/image/"+ qualificationsfordepartment;
            int[][] imagesPoint = new int[images.length][4];
            imagesPoint[0][0] = 1;
            imagesPoint[0][1] = 1;
            imagesPoint[0][2] = 9;
            imagesPoint[0][3] = 5;
            double [] scaleFactors = new double[images.length];
            scaleFactors[0] = 1.0;
            excelMthd.insertImages(localfmPath, 0, images, imagesPoint,scaleFactors);
            //添加条形码
            images = new String[1];
            images[0] = barcodePath + "/barcode.png";
            imagesPoint = new int[images.length][4];
            imagesPoint[0][0] = 2;
            imagesPoint[0][1] = 11;
            imagesPoint[0][2] = 5;
            imagesPoint[0][3] = 12;
            scaleFactors = new double[images.length];
            scaleFactors[0] = 0.7;
            excelMthd.insertImageIntoExcelCenter(localfmPath, 0, images, imagesPoint,scaleFactors);
            //添加二维码
            images = new String[1];
            images[0] = barcodePath + "/qrcode.png";
            imagesPoint = new int[images.length][4];
            imagesPoint[0][0] = 7;
            imagesPoint[0][1] = 11;
            imagesPoint[0][2] = 8;
            imagesPoint[0][3] = 13;
            scaleFactors = new double[images.length];
            scaleFactors[0] = 1.0;
            excelMthd.insertImagesCenter(localfmPath, 0, images, imagesPoint,scaleFactors);

            //拷贝到文件服务器
            String objectName = remotePath + "/fm.xlsx";
            minioUtil.uploadLocalFile(minioProperties.getBucketName(),localfmPath,objectName);

            createRespVO.setFmExcelPath(objectName);
        }
        //首页
        if (localsyPath != null && !localsyPath.isEmpty()) {
            int sheetIndex = 0;
            //填充数据
            excelMthd.reportStaticExcel(localsyPath, sheetIndex, syDataList);
            //拷贝到文件服务器
            String objectName = remotePath + "/sy.xlsx";
            minioUtil.uploadLocalFile(minioProperties.getBucketName(),localsyPath,objectName);

            createRespVO.setSyExcelPath(objectName);
        }
        //附页
        if ("自动".equals(reportgentype)){
            if (localfyPath != null && !localfyPath.isEmpty() && fyDataModel != null) {
                int sheetIndex = 0;
                //1:填充静态数据
                Map<String, String> fyStaticData = fyDataModel.getFYStaticData();
                excelMthd.reportStaticExcel(localfyPath, sheetIndex, fyStaticData);
                //2:填充表格数据
                excelMthd.reportFy(localfyPath, sheetIndex, fyDataModel, null);
                //拷贝到文件服务器
                String objectName = remotePath + "/fy.xlsx";
                minioUtil.uploadLocalFile(minioProperties.getBucketName(),localfyPath,objectName);

                createRespVO.setFyExcelPath(objectName);
            }
        }
        if ("合成".equals(reportgentype)){
            if (excelfyPath != null && !excelfyPath.isEmpty()){
                //拷贝到文件服务器
                String objectName = remotePath + "/fy.xlsx";
                minioUtil.uploadLocalFile(minioProperties.getBucketName(),excelfyPath,objectName);

                createRespVO.setFyExcelPath(objectName);
            }
        }
//        //合并pdf
//        //生成一个uuid字符串
//        String uuid = UUID.randomUUID().toString();
//        //添加注意事项pdf
//        String attachedPdf = localPath + "/attached_pdf/" + attachedPdfName;
//        pdfList.add(attachedPdf);
//        //合并pdf
//        utils.mergePdfFiles(pdfList, uploadPath + "/" + uuid + ".pdf");

//        List<String> listFiles = utils.listFiles(uploadPath);
//        CreateRespVO createRespVO = new CreateRespVO();
//        for (String filePath : listFiles) {
//            String fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
//            String remoteFilePath = remoteFtpPath + "/" + fileName;
//            System.out.println("filePath = " + filePath);
//            System.out.println("remoteFilePath = " + remoteFilePath);
//            boolean b = ftpUtil.uploadFile(ftpClient, filePath, remoteFilePath);
//            if (b) {
//                String fristPath = "ftp://@iphash@";
//                Map<String, Consumer<String>> fileNameToSetterMap = new HashMap<>();
//                fileNameToSetterMap.put("sy.pdf", createRespVO::setSyPdfPath);
//                fileNameToSetterMap.put("sy.xls", createRespVO::setSyExcelPath);
//                fileNameToSetterMap.put("fy.pdf", createRespVO::setFyPdfPath);
//                fileNameToSetterMap.put("fy.xls", createRespVO::setFyExcelPath);
//                fileNameToSetterMap.put("fm.pdf", createRespVO::setFmPdfPath);
//                fileNameToSetterMap.put("fm.xls", createRespVO::setFmExcelPath);
//                fileNameToSetterMap.put("fmtd.pdf", createRespVO::setFmtdPdfPath);
//                fileNameToSetterMap.put("fmtd.xls", createRespVO::setFmtdExcelPath);
//                fileNameToSetterMap.put(uuid + ".pdf", createRespVO::setPath2);
//
//                Consumer<String> setter = fileNameToSetterMap.get(fileName);
//                if (setter != null) {
//                    setter.accept(fristPath + remoteFilePath);
//                } else {
//                    // Handle the case where fileName doesn't match any key
//                }
//            }
//        }
//        ftpUtil.closeFTPClient(ftpClient);


        return success(createRespVO);
    }


    @PostMapping("/merge_pdf")
    @Operation(summary = "生成最终pdf")
    public CommonResult<WholePDFVO> createWholePdf(@Valid @RequestBody WholeExcel wholeExcel) {
        String folderno = wholeExcel.getFOLDERNO();
        String preordno = wholeExcel.getPREORDNO();
        String inspectiondate = wholeExcel.getINSPECTIONDATE();
        String qfDate = wholeExcel.getQFDATE();

        String excelFmSingleFtp = wholeExcel.getEXCEL_FM_SINGLE_FTP();
        String excelEditPath = wholeExcel.getEXCEL_EDIT_PATH();
        String excelFmFtp = wholeExcel.getEXCEL_FM_FTP();
        String excelSyFtp = wholeExcel.getEXCEL_SY_FTP();
        String excelFyFtp = wholeExcel.getEXCEL_FY_FTP();
        String excelFy2Path = wholeExcel.getEXCEL_FY2_PATH();
        String excelFy3Path = wholeExcel.getEXCEL_FY3_PATH();
        String attachedPdfName = wholeExcel.getAttached_PDF_Name();

        String stampBite = wholeExcel.getStampBite();
        String stampBiteQF = wholeExcel.getStampBiteQF();
        String sTestBy = wholeExcel.getSTestBy();
        String sApprovedBy = wholeExcel.getSApprovedBy();
        String sReleasedBy = wholeExcel.getSReleasedBy();

        //创建本地目录
        String localPath = utils.localpath;

        utils.createLocalDirectory(localPath);

        String stampPngPath = null;
        String stampQFPngPath = null;
        String sTestPngPath = null;
        String sApprovedPngPath = null;
        String sReleasedPngPath = null;
        //byte[]转图片
        if(stampBite != null && !stampBite.isEmpty()){
            stampPngPath = localPath + "/" + inspectiondate + "/" + folderno + "/stamp.png";
            byte[] stampBiteBytes = Base64.getDecoder().decode(stampBite);
            // 将字节数组转换为图片并保存
            utils.saveByteArrayAsPNG(stampBiteBytes, stampPngPath);
        }
        if(stampBiteQF != null && !stampBiteQF.isEmpty()){
            stampQFPngPath = localPath + "/" + inspectiondate + "/" + folderno + "/stamp_qf.png";
            byte[] stampBiteQFBytes = Base64.getDecoder().decode(stampBiteQF);
            // 将字节数组转换为图片并保存
            utils.saveByteArrayAsPNG(stampBiteQFBytes, stampQFPngPath);
        }
        if(sTestBy != null && !sTestBy.isEmpty()){
            sTestPngPath = localPath + "/" + inspectiondate + "/" + folderno + "/stestby.png";
            byte[] sTestByBytes = Base64.getDecoder().decode(sTestBy);
            // 将字节数组转换为图片并保存
            utils.saveByteArrayAsPNG(sTestByBytes, sTestPngPath);
        }
        if(sApprovedBy != null && !sApprovedBy.isEmpty()){
            sApprovedPngPath = localPath + "/" + inspectiondate + "/" + folderno + "/sapprovedby.png";
            byte[] sApprovedByBytes = Base64.getDecoder().decode(sApprovedBy);
            // 将字节数组转换为图片并保存
            utils.saveByteArrayAsPNG(sApprovedByBytes, sApprovedPngPath);
        }
        if(sReleasedBy != null && !sReleasedBy.isEmpty()){
            sReleasedPngPath = localPath + "/" + inspectiondate + "/" + folderno + "/sreleasedby.png";
            byte[] sReleasedByBytes = Base64.getDecoder().decode(sReleasedBy);
            // 将字节数组转换为图片并保存
            utils.saveByteArrayAsPNG(sReleasedByBytes, sReleasedPngPath);
        }

        String uploadPath = localPath+"/upload/"+preordno;

        utils.createLocalDirectory(uploadPath);

        String localfmtdPath = excelFmSingleFtp;//"./static/20231215/2374609/fmtd.xls";//null;
        String localfmPath = excelFmFtp;//"./static/20231215/2374609/fm.xls";//null;
        String localsyPath = null;//"./static/20231215/2374609/sy.xls";//null;
        String localfyPath = excelFyFtp;//"./static/20231215/2374609/fy.xls";//null;
//        String local2Path = null;
//        String local3Path = null;

//        //1:下载excel
//        //封页套打
//        if (excelFmSingleFtp != null && !excelFmSingleFtp.isEmpty()) {
//            localfmtdPath = localPath + excelFmSingleFtp;
//            minioUtil.downloadFile(minioProperties.getBucketName(),excelFmSingleFtp,localfmtdPath);
//        }
//
//        //封页
//        if (excelFmFtp != null && !excelFmFtp.isEmpty()) {
//            localfmPath = localPath + excelFmFtp;
//            minioUtil.downloadFile(minioProperties.getBucketName(),excelFmFtp,localfmPath);
//
//        }
        //首页
        if (excelSyFtp != null && !excelSyFtp.isEmpty()) {
            localsyPath = localPath + excelSyFtp;
            minioUtil.downloadFile(minioProperties.getBucketName(),excelSyFtp,localsyPath);

        }
//        //附页
//        if (excelFyFtp != null && !excelFyFtp.isEmpty()) {
//            localfyPath = localPath + excelFyFtp;
//            minioUtil.downloadFile(minioProperties.getBucketName(),excelFyFtp,localfyPath);
//        }

        String remotePath = "/" + inspectiondate + "/" + folderno;

        //2:转pdf
        String fmtdPdf = null;
        List<String> pdfList = new ArrayList<>();
        //List<String> pdfWaterMarkList = new ArrayList<>();
        List<String> mergePdfList = new ArrayList<>();
        //List<String> mergeQZPdfList = new ArrayList<>();
        //List<String> mergePdfWaterMarkList = new ArrayList<>();

        //封面套打
        if (localfmtdPath != null && !localfmtdPath.isEmpty()) {
            //拷贝到文件服务器
//            try {
//                utils.copyFile(localfmPath, uploadPath,"fmtd.xls");
//            } catch (IOException e) {
//                throw new RuntimeException(e);
//            }
            String objectName = remotePath + "/fmtd.xlsx";
//            minioUtil.uploadLocalFile(minioProperties.getBucketName(),localfmtdPath,objectName);

            String objectUrl = minioUtil.getPresignedObjectUrl(minioProperties.getBucketName(), objectName);

            // excel转pdf
            fmtdPdf = utils.excel2pdf("fmtd" ,uploadPath,objectUrl);
        }
        //封面
        if (localfmPath != null && !localfmPath.isEmpty()) {
            //拷贝到文件服务器
            String objectName = remotePath + "/fm.xlsx";
            //minioUtil.uploadLocalFile(minioProperties.getBucketName(),localfmPath,objectName);

            String objectUrl = minioUtil.getPresignedObjectUrl(minioProperties.getBucketName(), objectName);

            //excel转pdf
            String pdfPath = utils.excel2pdf("fm" ,uploadPath,objectUrl);
            if(pdfPath != null && !pdfPath.isEmpty()){
                pdfList.add(pdfPath);
            }
//            //5:加水印
//            if(stampPngPath != null && !stampPngPath.isEmpty()){
//                String pdfAddWater = utils.excel2pdfAddWater("fm", uploadPath,objectUrl);
//                if(pdfAddWater != null && !pdfAddWater.isEmpty()){
//                    pdfWaterMarkList.add(pdfAddWater);
//                }
//            }
        }
        //首页
        if (localsyPath != null && !localsyPath.isEmpty()) {
            //签发日期
            if (qfDate != null && !qfDate.isEmpty()){
                Map<String,String> qfData = new HashMap<>();
                qfData.put("&{签发日期}",qfDate);
                excelMthd.reportStaticExcel(localsyPath, 0, qfData);
            }
            //4:签名
            Map<String, String> signatures = new HashMap<>();
            if(sTestPngPath != null && !sTestPngPath.isEmpty()){
                signatures.put("&{主检}", sTestPngPath);
            }
            if(sApprovedPngPath != null && !sApprovedPngPath.isEmpty()){
                signatures.put("&{二审}", sApprovedPngPath);
            }
            if(sReleasedPngPath != null && !sReleasedPngPath.isEmpty()){
                signatures.put("&{终审}", sReleasedPngPath);
            }
            if(!signatures.isEmpty()){
                excelMthd.insertImages(localsyPath,0,signatures);
            }
            //utils.replaceTextWithImages(localsyPath, localsyPath, signatures);

            //拷贝到文件服务器
            String objectName = remotePath + "/sy.xlsx";
            minioUtil.uploadLocalFile(minioProperties.getBucketName(),localsyPath,objectName);

            String objectUrl = minioUtil.getPresignedObjectUrl(minioProperties.getBucketName(), objectName);

//            //5:加水印
//            if(stampPngPath != null && !stampPngPath.isEmpty()){
//                String pdfAddWater = utils.excel2pdfAddWater("sy", uploadPath,objectUrl);
//                if(pdfAddWater != null && !pdfAddWater.isEmpty()){
//                    mergePdfWaterMarkList.add(pdfAddWater);
//                }
//            }
            //6:签名excel转pdf
            String pdfPath = utils.excel2pdf("sy" ,uploadPath,objectUrl);
            if(pdfPath != null && !pdfPath.isEmpty()){
                mergePdfList.add(pdfPath);
            }

            //签章,签章不在这里实现，在生成pdf后再加签章
//            Map<String, String> qzs = new HashMap<>();
//            qzs.put("(检验检测专用章)", stampPngPath);
//            scale = excelMthd.insertImages(localsyPath,0,qzs);
//            //utils.replaceTextWithImages(localsyPath, localsyPath, qzs);
//            objectName = remotePath + "/sy.xlsx";
//            minioUtil.uploadLocalFile(minioProperties.getBucketName(),localsyPath,objectName);

//            objectUrl = minioUtil.getPresignedObjectUrl(minioProperties.getBucketName(), objectName);
//            System.out.println("objectUrl = " + objectUrl);
//            //6:签章excel转pdf
//            String qzpdfPath = utils.excel2pdf("sy" ,uploadPath,objectUrl);
//            if(qzpdfPath != null && !qzpdfPath.isEmpty()){
//                mergeQZPdfList.add(qzpdfPath);
//            }

        }
        //附页
        if (localfyPath != null && !localfyPath.isEmpty()) {

            String objectName = remotePath + "/fy.xlsx";
            //minioUtil.uploadLocalFile(minioProperties.getBucketName(),localfyPath,objectName);

            String objectUrl = minioUtil.getPresignedObjectUrl(minioProperties.getBucketName(), objectName);

            // excel转pdf
            String pdfPath = utils.excel2pdf("fy" ,uploadPath,objectUrl);
            if(pdfPath != null && !pdfPath.isEmpty()){

                mergePdfList.add(pdfPath);
                //mergeQZPdfList.add(pdfPath);
            }
//            //5:加水印
//            if(stampPngPath != null && !stampPngPath.isEmpty()){
//                String pdfAddWater = utils.excel2pdfAddWater("fy", uploadPath,objectUrl);
//                if(pdfAddWater != null && !pdfAddWater.isEmpty()){
//                    mergePdfWaterMarkList.add(pdfAddWater);
//                }
//            }

        }

        WholePDFVO wholePDFVO = new WholePDFVO();
        //生成一个uuid字符串
        String uuid = UUID.randomUUID().toString();
        //添加页码，有3个pdflist要添加页码
        String mergePDFPath = null;
        if(!mergePdfList.isEmpty()){
            String mergePdfPath = uploadPath + "/" + uuid + "_merge.pdf";
            utils.mergePDFAndAddPageNumbers(mergePdfList,mergePdfPath);
            mergePDFPath = mergePdfPath;
            pdfList.add(mergePdfPath);
        }
//        //3:合并pdf
//        if(!mergeQZPdfList.isEmpty()) {
//            String mergePdfPath = uploadPath + "/qz_merge.pdf";
//            utils.mergePDFAndAddPageNumbers(mergeQZPdfList,mergePdfPath);
//            pdfList.add(mergePdfPath);
//        }

//        if(!mergePdfWaterMarkList.isEmpty()) {
//            String mergePdfPath = uploadPath + "/watermark_merge.pdf";
//            utils.mergePDFAndAddPageNumbers(mergePdfWaterMarkList,mergePdfPath);
//            pdfWaterMarkList.add(mergePdfPath);
//        }
        //添加注意事项pdf
        if(attachedPdfName != null && !attachedPdfName.isEmpty()){
            String attachedPdf = localPath + "/attached_pdf/" + attachedPdfName;
            pdfList.add(attachedPdf);
//            if(stampPngPath != null && !stampPngPath.isEmpty()){
//                pdfWaterMarkList.add(attachedPdf);
//            }
        }
        String watermarkPdfName = null;
       if(pdfList.size() > 0) {
           //合并pdf
           utils.mergePdfFiles(pdfList, uploadPath + "/merge_print.pdf");
           String localPdfName = uploadPath + "/merge_print.pdf";

           float scale = 0.4f;
           if(stampPngPath != null && !stampPngPath.isEmpty()){
               watermarkPdfName = localPdfName;
               // 添加盖章
               utils.addStampQZToPDF(localPdfName,uploadPath + "/qz_print.pdf",stampPngPath,"(检验检测专用章)",1,1.0f, scale);
               localPdfName = uploadPath + "/qz_print.pdf";
           }
           if(stampQFPngPath != null && !stampQFPngPath.isEmpty()){
               //添加骑缝章
               utils.addStampToPDF(localPdfName,uploadPath + "/" + uuid + "_print.pdf",stampQFPngPath,0.7f, scale);
               localPdfName = uploadPath + "/" + uuid + "_print.pdf";
           }
            //编辑加密
           utils.PDFEditProtection(localPdfName,localPdfName);
           //拷贝到文件服务器
           String objectName = remotePath + "/" +  uuid + "_print.pdf";
           minioUtil.uploadLocalFile(minioProperties.getBucketName(),localPdfName,objectName);
           wholePDFVO.setPDF_PRINT_PATH(objectName);
       }else{
           wholePDFVO.setPDF_PRINT_PATH(null);
       }

       if(watermarkPdfName != null && !watermarkPdfName.isEmpty()) {
           String watermarkPdf = uploadPath + "/" + uuid + "_watermark.pdf";
           //添加水印
           utils.addWaterMarkToPDF(watermarkPdfName,watermarkPdf,"仅供预览 打印无效");
           //编辑加密
           utils.PDFEditProtection(watermarkPdf,watermarkPdf);

           String objectName = remotePath + "/" +  uuid + "_watermark.pdf";
           minioUtil.uploadLocalFile(minioProperties.getBucketName(),uploadPath + "/" + uuid + "_watermark.pdf",objectName);
           wholePDFVO.setPDF_WATERMARK_PATH(objectName);
       }else{
           wholePDFVO.setPDF_WATERMARK_PATH(null);
       }

        if (fmtdPdf != null && !fmtdPdf.isEmpty()) {
            String objectName = remotePath + "/fmtd.pdf";
            minioUtil.uploadLocalFile(minioProperties.getBucketName(), fmtdPdf, objectName);
            wholePDFVO.setPDF_FM_SINGLE_FTP(objectName);
        } else {
            wholePDFVO.setPDF_FM_SINGLE_FTP(null);
        }

        if (mergePDFPath != null && !mergePDFPath.isEmpty()) {
            String objectName = remotePath + "/" + uuid + "_merge.pdf";
            minioUtil.uploadLocalFile(minioProperties.getBucketName(), mergePDFPath, objectName);
            wholePDFVO.setMERGE_PATH(objectName);
        } else {
            wholePDFVO.setMERGE_PATH(null);
        }

        //上传一份到ftp
        FTPClient ftpClient = ftpUtil.initializeFTPClient();

        String fmtdPdfPath = remotePath + "/fmtd.pdf";
        String mergePdfPath = remotePath + "/" + uuid + "_merge.pdf";
        String printPdfPath = remotePath + "/" + uuid + "_print.pdf";
        String watermarkPdfPath = remotePath + "/" + uuid + "_watermark.pdf";

//        if(fmtdPdf != null && !fmtdPdf.isEmpty()){
//            ftpUtil.uploadFile(ftpClient, fmtdPdf, fmtdPdfPath);
//        }
        if (mergePDFPath != null && !mergePDFPath.isEmpty()) {
            ftpUtil.uploadFile(ftpClient, mergePDFPath, mergePdfPath);
        }
        if(stampPngPath != null && !stampPngPath.isEmpty()){
            ftpUtil.uploadFile(ftpClient, uploadPath + "/" + uuid + "_print.pdf", printPdfPath);
            ftpUtil.uploadFile(ftpClient, uploadPath + "/" + uuid + "_watermark.pdf", watermarkPdfPath);
        }

        ftpUtil.closeFTPClient(ftpClient);

        return success(wholePDFVO);
    }

    @PostMapping("/replace_image")
    @Operation(summary = "替换封面标头")
    public CommonResult<Boolean> replaceImage(@Valid @RequestBody ReplaceImage replace) {
        boolean result = false;
        String sourceExcel = replace.getSourceExcel();
        String pictureName = replace.getPictureName();
        String replaceImage = replace.getReplaceImage();
        log.info("pictureName = " + pictureName);

        String localExcelPath = null;
        String replaceImagePath = null;
        //创建本地目录
        String localPath = utils.localpath;
        utils.createLocalDirectory(localPath);

        if(sourceExcel != null && !sourceExcel.isEmpty()){
            localExcelPath = localPath + "/replace_excel/" + sourceExcel;
            minioUtil.downloadFile(minioProperties.getBucketName(),sourceExcel,localExcelPath);
        }
        if(replaceImage != null && !replaceImage.isEmpty()){
            replaceImagePath = localPath + "/image/" + replaceImage;
        }
        if(localExcelPath != null && replaceImagePath != null && pictureName != null && !pictureName.isEmpty()){
            boolean isRemove = excelMthd.removeExistingImage(localExcelPath,pictureName);
            if(isRemove){
                String[] images = new String[1];
                images[0] = replaceImagePath;
                int[][] imagesPoint = new int[images.length][4];
                imagesPoint[0][0] = 1;
                imagesPoint[0][1] = 1;
                imagesPoint[0][2] = 9;
                imagesPoint[0][3] = 5;
                double [] scaleFactors = new double[images.length];
                scaleFactors[0] = 1.0;
                excelMthd.insertImages(localExcelPath, 0, images, imagesPoint,scaleFactors);
                //上传回文件存储器
                String objectName = sourceExcel;
                minioUtil.uploadLocalFile(minioProperties.getBucketName(), localExcelPath, objectName);
                result = true;
            }
        }

        return success(result);
    }
}
