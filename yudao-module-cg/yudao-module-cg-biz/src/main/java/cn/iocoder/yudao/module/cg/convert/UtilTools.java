package cn.iocoder.yudao.module.cg.convert;

import com.alibaba.fastjson.JSONObject;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.AffineTransform;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.pdf.extgstate.PdfExtGState;
import com.itextpdf.kernel.utils.PdfMerger;
import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.PDExtendedGraphicsState;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.primeframework.jwt.Signer;
import org.primeframework.jwt.domain.JWT;
import org.primeframework.jwt.hmac.HMACSigner;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.file.*;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.List;
import java.util.*;

@Component
@Slf4j
public class UtilTools {

    @Value("${yudao.file-store-secret}")
    private String secret;

    @Value("${yudao.excel-pdf-url}")
    private String excelPdfUrl;

    @Value("${yudao.pdf-font-path}")
    private String pdfFontPath;

    @Value("${yudao.pdf-password}")
    private String pdfPassword;

    @Value("${yudao.static-path}")
    public String localpath;
    /***
     *
     * @param sourceFilePath：源文件路径
     * @param targetDirectoryPath：目标文件夹路径
     * @param newFileName：新文件名
     * @throws IOException
     */
    public void copyFile(String sourceFilePath, String targetDirectoryPath, String newFileName) throws IOException {
        Path sourceFile = Paths.get(sourceFilePath);
        Path targetDirectory = Paths.get(targetDirectoryPath);
        // 确保目标目录存在
        if (!Files.exists(targetDirectory)) {
            Files.createDirectories(targetDirectory);
        }
        // 目标文件路径
        Path targetFile = targetDirectory.resolve(newFileName);
        // 拷贝文件到目标目录并重命名
        Files.copy(sourceFile, targetFile, StandardCopyOption.REPLACE_EXISTING);

    }
    /**
     * 执行POST请求并获取返回值
     *
     * @param targetURL 目标URL
     * @param jsonInputString JSON格式的请求体
     * @return 服务器响应
     * @throws Exception 如果请求失败
     */
    public  String sendPostRequest(String targetURL, String jsonInputString){
        HttpURLConnection connection = null;

        try {
            // 创建URL对象
            URL url = new URL(targetURL);
            // 打开连接
            connection = (HttpURLConnection) url.openConnection();
            // 设置请求方法为POST
            connection.setRequestMethod("POST");
            // 设置请求头
            connection.setRequestProperty("Content-Type", "application/json; utf-8");
            connection.setRequestProperty("Accept", "application/json");
            // 允许写入请求体
            connection.setDoOutput(true);

            // 写入请求体
            try (OutputStream os = connection.getOutputStream()) {
                byte[] input = jsonInputString.getBytes("utf-8");
                os.write(input, 0, input.length);
            }

            // 获取响应码
            int responseCode = connection.getResponseCode();

            // 读取响应
            StringBuilder response = new StringBuilder();
            try (BufferedReader br = new BufferedReader(new InputStreamReader(connection.getInputStream(), "utf-8"))) {
                String responseLine;
                while ((responseLine = br.readLine()) != null) {
                    response.append(responseLine.trim());
                }
            }
            return response.toString();

        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
    }

    /***
     *
     * @param fileUrl：文件下载路径
     * @param localFilename：本地文件全路径，包含文件名
     * @throws IOException
     */
    public void downloadFile(String fileUrl, String localFilename) throws IOException {
        URL url = new URL(fileUrl);
        HttpURLConnection httpConn = (HttpURLConnection) url.openConnection();
        int responseCode = httpConn.getResponseCode();

        // 检查 HTTP 响应代码
        if (responseCode == HttpURLConnection.HTTP_OK) {
            try (InputStream inputStream = httpConn.getInputStream();
                 BufferedInputStream in = new BufferedInputStream(inputStream);
                 FileOutputStream fileOutputStream = new FileOutputStream(localFilename)) {

                byte[] buffer = new byte[4096];
                int bytesRead;
                while ((bytesRead = in.read(buffer)) != -1) {
                    fileOutputStream.write(buffer, 0, bytesRead);
                }
            }
        } else {
            throw new IOException("No file to download. Server replied with: " + responseCode);
        }
        httpConn.disconnect();
    }
    public  String createToken(final Map<String, Object> payloadClaims) {
        try {
            // build a HMAC signer using a SHA-256 hash
            //secret=${files.docservice.secret}
            //zXS4A5E9BR5FU2adqrfZ3yj12PogCG
            Signer signer = HMACSigner.newSHA256Signer(secret);
            JWT jwt = new JWT();
            for (String key : payloadClaims.keySet()) {  // run through all the keys from the payload
                jwt.addClaim(key, payloadClaims.get(key));  // and write each claim to the jwt
            }
            return JWT.getEncoder().encode(jwt, signer);  // sign and encode the JWT to a JSON string representation
        } catch (Exception e) {
            return "";
        }
    }

    public String getMD5Hash(String input) {
        // 创建一个MD5算法实例
        MessageDigest md = null;
        try {
            md = MessageDigest.getInstance("MD5");
        } catch (NoSuchAlgorithmException e) {
            throw new RuntimeException(e);
        }
        // 更新摘要
        md.update(input.getBytes());
        // 计算哈希值
        byte[] digest = md.digest();
        // 将字节数组转换为十六进制字符串并返回
        return bytesToHex(digest);
    }

    // 将字节数组转换为十六进制字符串
    public String bytesToHex(byte[] bytes) {
        StringBuilder sb = new StringBuilder();
        for (byte b : bytes) {
            sb.append(String.format("%02x", b));
        }
        return sb.toString();
    }

    public String excel2pdf(String fileName,String uploadPath,String objectUrl) {
        Map<String,Object> margin = new HashMap<>();
        margin.put("left","15mm");
        margin.put("right","5mm");
        margin.put("top","19.1mm");
        margin.put("bottom","10mm");
        Map<String,Object> fit = new HashMap<>();
        fit.put("scale", 100);
        fit.put("margins",margin);
        int randomNum = (int) (Math.random() * 100);
        String key = getMD5Hash(objectUrl) + randomNum;
        System.out.println("key = " + key);
        Map<String, Object> map = new HashMap<>();
        map.put("async", false);
        map.put("filetype", "xlsx");
        map.put("key", key);
        map.put("outputtype","pdf");
        map.put("title",fileName);
        map.put("url",objectUrl);
        map.put("spreadsheetLayout",fit);

        String token =  createToken(map);
        map.put("token", token);
        //map转为json
        String json = JSONObject.toJSONString(map);

        String result = sendPostRequest(excelPdfUrl, json);
        // 解析 JSON
        JSONObject jsonResponse = JSONObject.parseObject(result);

        // 检查 endConvert 是否为 true
        Boolean endConvert = jsonResponse.getBoolean("endConvert");
        if (Boolean.TRUE.equals(endConvert)) {
            // 获取 fileUrl
            String fileUrl = jsonResponse.getString("fileUrl");
            // 下载文件
            try {
                downloadFile(fileUrl, uploadPath + "/" + fileName + ".pdf");
                return uploadPath + "/" + fileName + ".pdf";
            } catch (IOException e) {
                System.err.println("Failed to download file: " + e.getMessage());
            }
        }
        return null;
    }
    public String excel2pdfAddWater(String fileName,String uploadPath,String objectUrl) {

        Map<String, Object> wartmark = new HashMap<>();
        wartmark.put("transparent", 0.4);
        wartmark.put("type", "rect");
        wartmark.put("width", 300);
        wartmark.put("height", 50);
        wartmark.put("rotate", 45);
        List<Integer> margins = new ArrayList<>();
        margins.add(10);
        margins.add(10);
        margins.add(10);
        margins.add(10);
        wartmark.put("margins", margins);
        wartmark.put("align", 1);
        List<Map<String, Object>> paragraphs = new ArrayList<>();
        Map<String, Object> paragraph = new HashMap<>();
        paragraph.put("align", 2);
        paragraph.put("linespacing", 1);
        List<Map<String, Object>> runs = new ArrayList<>();
        Map<String, Object> run = new HashMap<>();
        run.put("text", "仅供预览 打印无效");
        List<Integer> fill = new ArrayList<>();
        fill.add(204);
        fill.add(204);
        fill.add(204);
        run.put("fill", fill);
        run.put("font-family", "Arial");
        run.put("font-size", 72);
        run.put("bold", true);
        run.put("italic", false);
        run.put("strikeout", false);
        run.put("underline", false);
        runs.add(run);
        paragraph.put("runs", runs);
        paragraphs.add(paragraph);

        wartmark.put("paragraphs", paragraphs);

        Map<String,Object> margin = new HashMap<>();
        margin.put("left","15mm");
        margin.put("right","5mm");
        margin.put("top","19.1mm");
        margin.put("bottom","10mm");
        Map<String,Object> fit = new HashMap<>();
        fit.put("scale", 100);
        fit.put("margins",margin);
        int randomNum = (int) (Math.random() * 100);
        String key = getMD5Hash(objectUrl) + randomNum;
        System.out.println("key = " + key);
        Map<String, Object> map = new HashMap<>();
        map.put("async", false);
        map.put("filetype", "xlsx");
        map.put("key", key);
        map.put("outputtype","pdf");
        map.put("title",fileName);
        map.put("url", objectUrl);
        map.put("spreadsheetLayout",fit);
        map.put("watermark",wartmark);
        String token =  createToken(map);
        map.put("token", token);
        //map转为json
        String json = JSONObject.toJSONString(map);

        String result = sendPostRequest(excelPdfUrl, json);
        // 解析 JSON
        JSONObject jsonResponse = JSONObject.parseObject(result);

        // 检查 endConvert 是否为 true
        Boolean endConvert = jsonResponse.getBoolean("endConvert");
        if (Boolean.TRUE.equals(endConvert)) {
            // 获取 fileUrl
            String fileUrl = jsonResponse.getString("fileUrl");
            // 下载文件
            try {
                downloadFile(fileUrl, uploadPath + "/" + fileName + "_watermark.pdf");
                return uploadPath + "/" + fileName + "_watermark.pdf";
            } catch (IOException e) {
                System.err.println("Failed to download file: " + e.getMessage());
            }
        }
        return null;
    }

    public void mergePdfFiles(List<String> srcFiles, String dest) {
        // 创建一个 PDF 文档用于合并后的输出
        PdfDocument pdfDoc = null;
        try {
            pdfDoc = new PdfDocument(new PdfWriter(dest));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        PdfMerger merger = new PdfMerger(pdfDoc);

        for (String src : srcFiles) {
            // 打开每个源 PDF 文件
            PdfDocument srcDoc = null;
            try {
                srcDoc = new PdfDocument(new PdfReader(src));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            // 合并源 PDF 文件到输出 PDF 文档
            merger.merge(srcDoc, 1, srcDoc.getNumberOfPages());
            // 关闭源 PDF 文档
            srcDoc.close();
        }

        // 关闭输出 PDF 文档
        pdfDoc.close();
    }

    /***
     *
     * @param pdfsPath: 要合并的 PDF 文件路径
     * @param mergedPdfPath: 合并后的 PDF 文件路径
     */
    public void mergePDFAndAddPageNumbers(List<String> pdfsPath, String mergedPdfPath) {

        // 合并 PDF 文件
        PDFMergerUtility pdfMerger = new PDFMergerUtility();
        try {
            for(String pdfPath:pdfsPath){
                pdfMerger.addSource(new File(pdfPath));
            }
            pdfMerger.setDestinationFileName(mergedPdfPath);
            pdfMerger.mergeDocuments(null);
            // 打开合并后的 PDF 文件
            PDDocument mergedDocument = PDDocument.load(new File(mergedPdfPath));

            // 加载 TrueType 字体文件
            PDType0Font font = PDType0Font.load(mergedDocument, new File(pdfFontPath));

            // 获取总页数
            int totalPages = mergedDocument.getNumberOfPages();

            // 添加页码和自定义文本
            for (int i = 0; i < totalPages; i++) {
                PDPage page = mergedDocument.getPage(i);

                float reportNumberY = findTextPosition(mergedDocument, i, "报告编号");

                if (reportNumberY == -1) {
                    System.out.println("未找到“报告编号”文本");
                    continue;
                }

                PDPageContentStream contentStream = new PDPageContentStream(mergedDocument, page, PDPageContentStream.AppendMode.APPEND, true, true);

                // 设置字体和大小
                contentStream.setFont(font, 10);

                // 构建页码和自定义文本
                String text = "共" + totalPages + "页第" + (i + 1) + "页 Page No:" + totalPages + "-" + (i + 1);

                // 获取页面宽度
                float pageWidth = page.getMediaBox().getWidth();

                // 计算文本位置（右上角）
                float textWidth = (float) (font.getStringWidth(text) / 1000 * 12);

                float textX = pageWidth - textWidth - 50 + 20; // 调整文本位置
                float textY = reportNumberY; // 将文本 y 坐标设置为与“报告编号”相同

                // 添加文本到页面
                contentStream.beginText();
                contentStream.newLineAtOffset(textX, textY);
                contentStream.showText(text);
                contentStream.endText();

                contentStream.close();
            }
            // 保存并关闭文档
            mergedDocument.save(mergedPdfPath);
            mergedDocument.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    // 查找每页中“报告编号”文本位置的辅助方法
    // Modified findTextPosition method
    public float findTextPosition(PDDocument document, int pageNumber, String searchText) throws IOException {
        class MyPDFTextStripper extends PDFTextStripper {
            private float y = -1;
            private boolean found = false;

            public MyPDFTextStripper() throws IOException {}

            @Override
            protected void writeString(String str, List<TextPosition> textPositions) {
                StringBuilder sb = new StringBuilder();
                for (TextPosition text : textPositions) {
                    sb.append(text.getUnicode());
                    if (sb.toString().contains(searchText) && !found) {
                        y = text.getTextMatrix().getTranslateY();  // Using text matrix to capture precise baseline Y position
                        found = true;  // Ensures only the first occurrence is used
                    }
                }
            }

            public float getYPosition() {
                return y;
            }
        }

        MyPDFTextStripper stripper = new MyPDFTextStripper();
        stripper.setSortByPosition(true);
        stripper.setStartPage(pageNumber + 1);
        stripper.setEndPage(pageNumber + 1);
        stripper.getText(document);
        return stripper.getYPosition();
    }

    /***
     * 添加骑缝章
     * @param inputPdfPath：输入PDF
     * @param outputPdfPath：输出PDF
     * @param stampImagePath：印章图片路径
     * @param opacity：透明度：0.7f
     * @throws IOException
     */
    public void addStampToPDF(String inputPdfPath, String outputPdfPath, String stampImagePath, float opacity) {

        try (PDDocument document = PDDocument.load(new File(inputPdfPath))) {
            int pageCount = document.getNumberOfPages();
            BufferedImage stampImage = ImageIO.read(new File(stampImagePath));

            int stampHeight = stampImage.getHeight();
            int stampWidth = stampImage.getWidth();
            int singleStampWidth = stampWidth / pageCount;

            for (int i = 0; i < pageCount; i++) {
                PDPage page = document.getPage(i);
                BufferedImage subImage = stampImage.getSubimage(i * singleStampWidth, 0, singleStampWidth, stampHeight);

                // 创建一个具有 Alpha 通道的新图像
                BufferedImage transparentImage = new BufferedImage(singleStampWidth, stampHeight, BufferedImage.TYPE_INT_ARGB);
                Graphics2D g = transparentImage.createGraphics();
                g.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, opacity));
                g.drawImage(subImage, 0, 0, null);
                g.dispose();

                File tempFile = File.createTempFile("stamp", ".png");
                ImageIO.write(transparentImage, "png", tempFile);

                PDImageXObject pdImage = PDImageXObject.createFromFile(tempFile.getAbsolutePath(), document);

                try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
                    float pageHeight = page.getMediaBox().getHeight();
                    float pageWidth = page.getMediaBox().getWidth();
                    float yPosition = (pageHeight - stampHeight) / 2; // 计算垂直居中位置

                    // 设置图片的透明度
                    PDExtendedGraphicsState graphicsState = new PDExtendedGraphicsState();
                    graphicsState.setNonStrokingAlphaConstant(opacity);
                    contentStream.setGraphicsStateParameters(graphicsState);

                    contentStream.drawImage(pdImage, pageWidth - singleStampWidth, yPosition, singleStampWidth, stampHeight);
                }

                tempFile.delete();
            }

            document.save(new File(outputPdfPath));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void addStampToPDF(String inputPdfPath, String outputPdfPath, String stampImagePath, float opacity, float scale) {

        try (PDDocument document = PDDocument.load(new File(inputPdfPath))) {
            int pageCount = document.getNumberOfPages();
            BufferedImage stampImage = ImageIO.read(new File(stampImagePath));

            int stampHeight = stampImage.getHeight();
            int stampWidth = stampImage.getWidth();
            int singleStampWidth = stampWidth / pageCount;

            for (int i = 0; i < pageCount; i++) {
                PDPage page = document.getPage(i);
                BufferedImage subImage = stampImage.getSubimage(i * singleStampWidth, 0, singleStampWidth, stampHeight);

                // 对透明图像进行缩放，并设置透明度
                int scaledWidth = (int) (singleStampWidth * scale);
                int scaledHeight = (int) (stampHeight * scale);
                BufferedImage scaledImage = new BufferedImage(scaledWidth, scaledHeight, BufferedImage.TYPE_INT_ARGB);
                Graphics2D g2d = scaledImage.createGraphics();
                g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, opacity));
                g2d.drawImage(subImage, 0, 0, scaledWidth, scaledHeight, null);
                g2d.dispose();

                File tempFile = File.createTempFile("stamp", ".png");
                ImageIO.write(scaledImage, "png", tempFile);

                PDImageXObject pdImage = PDImageXObject.createFromFile(tempFile.getAbsolutePath(), document);

                try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
                    float pageHeight = page.getMediaBox().getHeight();
                    float pageWidth = page.getMediaBox().getWidth();
                    float yPosition = (pageHeight - scaledHeight) / 2; // 计算垂直居中位置

                    contentStream.drawImage(pdImage, pageWidth - scaledWidth, yPosition, scaledWidth, scaledHeight);
                }

                tempFile.delete();
            }

            document.save(new File(outputPdfPath));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public List<String> listFiles(String directoryPath) {
        List<String> list = new ArrayList<>();
        Path dirPath = Paths.get(directoryPath);

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dirPath)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    // 如果是目录，递归调用
                    listFiles(path.toString());
                } else {
                    // 如果是文件，打印文件路径
                    System.out.println(path.toString());
                    list.add(path.toString());
                }
            }
        } catch (IOException | DirectoryIteratorException e) {
            e.printStackTrace();
        }
        return list;
    }

    // 将字节数组保存为图片文件
    public void saveByteArrayAsPNG(byte[] imageBytes, String outputPath) {
        try (ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes)) {
            BufferedImage image = ImageIO.read(bis);
            if (image != null) {
                File outputFile = new File(outputPath);
                File parentDir = outputFile.getParentFile();
                // 检查并创建输出路径的父目录
                if (parentDir != null && !parentDir.exists()) {
                    if (!parentDir.mkdirs()) {
                        System.out.println("Failed to create directory: " + parentDir.getAbsolutePath());
                        return;
                    }
                }
                String formatName = outputPath.substring(outputPath.lastIndexOf('.') + 1);
                boolean result = ImageIO.write(image, formatName, outputFile);
                if (!result) {
                    System.out.println("Failed to save image as " + formatName);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    // 将字节数组保存为 JPEG 格式图片文件
    // 将字节数组保存为 JPEG 格式图
    public void saveByteArrayAsJPEG(byte[] imageBytes, String outputPath) throws IOException {
        ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);
        BufferedImage image = ImageIO.read(bis);
        if (image != null) {
            // 如果图像包含透明度，转换为 RGB 格式
            if (image.getColorModel().hasAlpha()) {
                image = convertToRGB(image);
            }

            File outputFile = new File(outputPath);
            File parentDir = outputFile.getParentFile();
            // 检查并创建输出路径的父目录
            if (parentDir != null && !parentDir.exists()) {
                if (!parentDir.mkdirs()) {
                    System.out.println("Failed to create directory: " + parentDir.getAbsolutePath());
                    return;
                }
            }
            try (ImageOutputStream ios = ImageIO.createImageOutputStream(outputFile)) {
                // 获取 JPEG 格式的 ImageWriter
                Iterator<ImageWriter> writers = ImageIO.getImageWritersByFormatName("jpeg");
                if (!writers.hasNext()) {
                    throw new IllegalStateException("No writers found for format: jpeg");
                }
                ImageWriter writer = writers.next();
                writer.setOutput(ios);
                writer.write(image);
            }
        }
    }

    private BufferedImage convertToRGB(BufferedImage image) {
        BufferedImage rgbImage = new BufferedImage(image.getWidth(), image.getHeight(), BufferedImage.TYPE_INT_RGB);
        Graphics2D g = rgbImage.createGraphics();
        g.drawImage(image, 0, 0, null);
        g.dispose();
        return rgbImage;
    }

    public  void replaceTextWithImages(String excelFilePath, String outputFilePath, Map<String, String> signatures) {
        try (InputStream inp = new FileInputStream(excelFilePath);
             Workbook workbook = WorkbookFactory.create(inp)) {

            Sheet sheet = workbook.getSheetAt(0);  // Assuming the data is in the first sheet

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (signatures.containsKey(cellValue)) {
                            String imagePath = signatures.get(cellValue);
                            CellRangeAddress mergedRegion = getMergedRegion(sheet, cell.getRowIndex(), cell.getColumnIndex());
                            if (cellValue.equals("(检验检测专用章)")){
                                insertImageQZ(workbook, sheet, mergedRegion, imagePath);
                            }else{
                                insertImage(workbook, sheet, mergedRegion, imagePath);
                            }
                            //                          insertImage(workbook, sheet, cell.getRowIndex(), cell.getColumnIndex(), imagePath);
                            cell.setCellValue("");  // Clear the cell value
                        }
                    }
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    private  CellRangeAddress getMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return new CellRangeAddress(rowIndex, rowIndex, colIndex, colIndex);
    }
    private void insertImage(Workbook workbook, Sheet sheet, CellRangeAddress region, String imagePath) throws IOException {
        InputStream inputStream = new FileInputStream(imagePath);
        byte[] bytes = inputStream.readAllBytes();
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        inputStream.close();

        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(region.getFirstColumn());
        anchor.setRow1(region.getFirstRow());
        anchor.setCol2(region.getLastColumn() + 1); // 确保覆盖整个区域
        anchor.setRow2(region.getLastRow() + 1);   // 确保覆盖整个区域
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

        Picture pict = drawing.createPicture(anchor, pictureIdx);
        pict.resize();

        // Calculate the total width and height of the merged region
        float totalWidth = 0;
        for (int col = region.getFirstColumn(); col <= region.getLastColumn(); col++) {
            totalWidth += sheet.getColumnWidthInPixels(col);
        }

        float totalHeight = 0;
        for (int row = region.getFirstRow(); row <= region.getLastRow(); row++) {
            totalHeight += sheet.getRow(row).getHeightInPoints() / 72 * 96;  // Convert points to pixels
        }
        double width = pict.getImageDimension().getWidth();
        double height = pict.getImageDimension().getHeight();
        // Resize picture to fit merged cell
        double scaleX = totalWidth / width;
        double scaleY = totalHeight / height;
        double scale = Math.min(scaleX, scaleY);
        pict.resize(scale);
    }
    private static void insertImageQZ(Workbook workbook, Sheet sheet, CellRangeAddress region, String imagePath) throws IOException {
        InputStream inputStream = new FileInputStream(imagePath);
        byte[] bytes = inputStream.readAllBytes();
        inputStream.close();
// 获取画图管理器
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper helper = workbook.getCreationHelper();

        BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imagePath));
        //图片的真实长宽
        int imageWidth = bufferedImage.getWidth();
        int imageHeight = bufferedImage.getHeight();
        // 获取图片起始单元格
        int col1 = region.getFirstColumn();
        int row1 =  region.getFirstRow() - 4 ;
        // 计算单元格的长宽
        float cellWidth = sheet.getColumnWidthInPixels(col1);
        float cellHeight = sheet.getRow(row1).getHeightInPoints() / 72 * 96; // 高度转换为像素
        // 计算需要的长宽比例的系数
        double a = imageWidth * 0.6 / cellWidth;
        double b = imageHeight * 0.5 / cellHeight;
        ClientAnchor anchor = helper.createClientAnchor();

        anchor.setCol1(col1);
        anchor.setRow1(row1);

        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        Picture pict = drawing.createPicture(anchor, pictureIdx);
        pict.resize(a,b);
        //printAnchorValues(anchor);
        //pict.resize(scale);
    }

    //写个方法：创建本地目录，如果存在就跳过，如果不存在则创建
    public boolean createLocalDirectory(String localpath) {
        boolean created = false;
        File directory = new File(localpath);
        if (!directory.exists()) {
            created = directory.mkdirs();
            if (created) {
                log.info("Directory created: " + localpath);
            } else {
                log.error("Failed to create directory: " + localpath);
            }
        }
        return created; // Directory was created successfully or failed to create
    }
    public  void addStampQZToPDF(String inputPDF, String outputPDF, String stampImage, String keyword,int pageNum,float opacity, float scale) {
        try (PDDocument document = PDDocument.load(new File(inputPDF))) {
            PDPage page = document.getPage(pageNum);

            //String keyword = "(检验检测专用章)";
            float xPosition = 400;
            float yPosition = 100;

            // 获取关键字的位置
            Optional<TextPosition> keywordPosition = findFirstKeywordPosition(document, keyword, pageNum);
            if (keywordPosition.isPresent()) {
                TextPosition firstPosition = keywordPosition.get();
                xPosition = firstPosition.getXDirAdj();
                //yPosition = firstPosition.getEndY();
                yPosition = firstPosition.getEndY() - (firstPosition.getHeightDir() + firstPosition.getFontSizeInPt());
            } else {
                System.out.println("关键字未找到，使用默认位置");
            }

            // 添加图章
            PDImageXObject stamp = PDImageXObject.createFromFile(stampImage, document);
            float stampWidth = stamp.getWidth() * scale;
            float stampHeight = stamp.getHeight() * scale;

            // 设置图章的透明度
            try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
                contentStream.setGraphicsStateParameters(createGraphicsState(document, opacity));
                contentStream.drawImage(stamp, xPosition, yPosition, stampWidth, stampHeight);
            }
            document.save(outputPDF);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Optional<TextPosition> findFirstKeywordPosition(PDDocument document, String keyword, int pageNum) throws IOException {
        List<TextPosition> textPositions = new ArrayList<>(); // 定义 textPositions
        PDFTextStripper stripper = new PDFTextStripper() {
            @Override
            protected void writeString(String string, List<TextPosition> positions) throws IOException {
                int index = string.indexOf(keyword);
                if (index != -1) {
                    textPositions.addAll(positions.subList(index, index + keyword.length()));
                }
            }
        };
        stripper.setSortByPosition(true);
        stripper.setStartPage(pageNum + 1);
        stripper.setEndPage(pageNum + 1);
        stripper.getText(document);

        return textPositions.isEmpty() ? Optional.empty() : Optional.of(textPositions.get(0));
    }


    private  PDExtendedGraphicsState createGraphicsState(PDDocument document, float opacity) {
        PDExtendedGraphicsState graphicsState = new PDExtendedGraphicsState();
        graphicsState.setNonStrokingAlphaConstant(opacity);
        return graphicsState;
    }

    //加密
    public void PDFEditProtection(String inputPDF,String outputPDF){
        try (PDDocument document = PDDocument.load(new File(inputPDF))) {
            // 设置权限，允许查看但不允许编辑
            AccessPermission accessPermission = new AccessPermission();
            accessPermission.setCanModify(false); // 禁止编辑
            accessPermission.setCanExtractContent(false); // 禁止提取内容
            accessPermission.setCanFillInForm(false); // 禁止填写表单
            accessPermission.setCanPrint(true); // 禁止打印

            // 设置加密策略，没有设置用户密码，只设置拥有者密码
            StandardProtectionPolicy policy = new StandardProtectionPolicy(pdfPassword, null, accessPermission);
            policy.setEncryptionKeyLength(128); // 设置加密密钥长度，128 或 256 位
            policy.setPermissions(accessPermission);

            // 应用加密保护
            document.protect(policy);
            // 保存加密后的 PDF 文件
            document.save(outputPDF);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void addWaterMarkToPDF(String src, String dest,String watermark) {
        try(PdfDocument pdfDoc = new PdfDocument(new PdfReader(src), new PdfWriter(dest))) {
            int n = pdfDoc.getNumberOfPages();

            // 加载支持中文的字体（例如SimSun）
            PdfFont font = PdfFontFactory.createFont("STSongStd-Light", "UniGB-UCS2-H");

            for (int i = 1; i <= n; i++) {
                PdfPage page = pdfDoc.getPage(i);
                com.itextpdf.kernel.geom.Rectangle pageSize = page.getPageSizeWithRotation();

                // 获取页面的内容流以确保水印位于最上层
                PdfCanvas canvas = new PdfCanvas(page);

                // 设置透明度
                PdfExtGState gs1 = new PdfExtGState().setFillOpacity(0.1f);
                canvas.setExtGState(gs1);

                // 设置水印的字体和大小
                canvas.beginText();
                canvas.setFontAndSize(font, 80);

                // 设置旋转角度和文本位置，确保水印以-45度角斜着放置
                AffineTransform transform = AffineTransform.getRotateInstance(Math.toRadians(-45),
                        pageSize.getWidth() - 200, pageSize.getHeight() - 50); // 设置旋转中心为页面的右上角附近
                canvas.setTextMatrix(transform);

                // 从指定的xPosition和yPosition添加水印
                float xPosition = pageSize.getWidth() / 2; // 调整x轴位置
                float yPosition = pageSize.getHeight() / 2; // 调整y轴位置
                canvas.moveText(xPosition, yPosition);
                canvas.showText(watermark);
            }
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        // 读取Excel文件
        InputStream fis = null;
        try {
            fis = new FileInputStream("/Users/chengang/Documents/ss.xlsx");
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(fis);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Sheet sheet =  workbook.getSheetAt(0);

    }
}
