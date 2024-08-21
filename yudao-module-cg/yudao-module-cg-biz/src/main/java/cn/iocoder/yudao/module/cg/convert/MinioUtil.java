package cn.iocoder.yudao.module.cg.convert;

import cn.hutool.core.io.IoUtil;
import io.minio.*;
import io.minio.http.Method;
import io.minio.messages.DeleteError;
import io.minio.messages.DeleteObject;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.tika.Tika;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.*;
import java.nio.file.Files;
import java.util.List;
import java.util.concurrent.TimeUnit;

@Slf4j
@Component
public class MinioUtil {
    @Autowired
    private MinioClient minioClient;

//    /**
//     * 文件上传/文件分块上传
//     *
//     * @param bucketName 桶名称
//     * @param objectName 对象名称
//     * @param sliceIndex 分片索引
//     * @param file       文件
//     */
//    public Boolean uploadFile(String bucketName, String objectName, Integer sliceIndex, MultipartFile file) {
//        try {
//            if (sliceIndex != null) {
//                objectName = objectName.concat("/").concat(Integer.toString(sliceIndex));
//            }
//            // 写入文件
//            minioClient.putObject(PutObjectArgs.builder()
//                    .bucket(bucketName)
//                    .object(objectName)
//                    .stream(file.getInputStream(), file.getSize(), -1)
//                    .contentType(file.getContentType())
//                    .build());
//            log.debug("上传到minio文件|uploadFile|参数：bucketName：{}，objectName：{}，sliceIndex：{}"
//                    , bucketName, objectName, sliceIndex);
//            return true;
//        } catch (Exception e) {
//            log.error("文件上传到Minio异常|参数：bucketName:{},objectName:{},sliceIndex:{}|异常:{}", bucketName, objectName, sliceIndex, e);
//            return false;
//        }
//    }

    /**
     * 创建桶，放文件使用
     *
     * @param bucketName 桶名称
     */
    public Boolean createBucket(String bucketName) {
        try {
            if (!minioClient.bucketExists(
                    BucketExistsArgs.builder().bucket(bucketName).build())) {
                minioClient.makeBucket(MakeBucketArgs.builder().bucket(bucketName).build());
            }
            return true;
        } catch (Exception e) {
            log.error("Minio创建桶异常!|参数：bucketName:{}|异常:{}", bucketName, e);
            return false;
        }
    }

    /**
     * 文件合并
     *
     * @param bucketName       桶名称
     * @param objectName       对象名称
     * @param sourceObjectList 源文件分片数据
     */
    public Boolean composeFile(String bucketName, String objectName, List<ComposeSource> sourceObjectList) {
        // 合并操作
        try {
            ObjectWriteResponse response = minioClient.composeObject(
                    ComposeObjectArgs.builder()
                            .bucket(bucketName)
                            .object(objectName)
                            .sources(sourceObjectList)
                            .build());
            return true;
        } catch (Exception e) {
            log.error("Minio文件按合并异常!|参数：bucketName:{},objectName:{}|异常:{}", bucketName, objectName, e);
            return false;
        }
    }

    /**
     * 多个文件删除
     *
     * @param bucketName 桶名称
     */
    public Boolean removeFiles(String bucketName, List<DeleteObject> delObjects) {
        try {
            Iterable<Result<DeleteError>> results =
                    minioClient.removeObjects(
                            RemoveObjectsArgs.builder().bucket(bucketName).objects(delObjects).build());
            boolean isFlag = true;
            for (Result<DeleteError> result : results) {
                DeleteError error = result.get();
                log.error("Error in deleting object {} | {}", error.objectName(), error.message());
                isFlag = false;
            }
            return isFlag;
        } catch (Exception e) {
            log.error("Minio多个文件删除异常!|参数：bucketName:{},objectName:{}|异常:{}", bucketName, e);
            return false;
        }
    }

    public boolean isBucketExist(String bucketName) {
        // Set the builder
        BucketExistsArgs.Builder builder = BucketExistsArgs.builder();
        builder.bucket(bucketName);

        BucketExistsArgs bucketExistsArgs = builder.build();

        boolean isExist = false;
        try {
            isExist = minioClient.bucketExists(bucketExistsArgs);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (isExist) {
            System.out.println("The bucket is exist! ");
            return true;
        } else {
            try {
                minioClient.makeBucket(MakeBucketArgs.builder().bucket(bucketName).build());
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            System.out.println("Create bucket is success! ");
            return true;
        }
    }

    @SneakyThrows
    public String getPresignedObjectUrl(String bucketName, String objectName){
        String uploadUrl = minioClient.getPresignedObjectUrl(GetPresignedObjectUrlArgs.builder()
                .method(Method.GET)
                .bucket(bucketName)
                .object(objectName)
                .expiry(10, TimeUnit.MINUTES) // 过期时间（秒数）取值范围：1 秒 ~ 7 天
                .build()
        );
        return uploadUrl;
    }

    public void uploadLocalFile(String bucketName,String localFilePath,String objectPath) {
        File localFile = new File(localFilePath);
        byte[] content = new byte[(int) localFile.length()];

        try (FileInputStream fis = new FileInputStream(localFile)) {
            fis.read(content);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        String mimeType = null;
        try {
            mimeType = Files.probeContentType(localFile.toPath());
            if (mimeType == null) {
                Tika tika = new Tika();
                mimeType = tika.detect(localFile);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        System.out.println("localFile.toPath()=="+localFile.toPath()+",mimeType = " + mimeType);
//        String objectPath = localFile.getName(); // 使用文件名作为对象路径
//        objectPath = objectPrefix + "/" + objectPath;
        System.out.println("objectPath = " + objectPath);
        try {
            minioClient.putObject(PutObjectArgs.builder()
                    .bucket(bucketName) // bucket 必须传递
                    .contentType(mimeType) // 使用确定的 MIME 类型
                    .object(objectPath) // 相对路径作为 key
                    .stream(new ByteArrayInputStream(content), content.length, -1) // 文件内容
                    .build());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void downloadFile(String bucketName,String objectName,String outputFilePath) {
        GetObjectResponse response = null;
        try {
            response = minioClient.getObject(GetObjectArgs.builder()
                    .bucket(bucketName) // bucket 必须传递
                    .object(objectName) // 相对路径作为 key
                    .build());
            byte[] bytes = IoUtil.readBytes(response);
            saveBytesToFile(bytes, outputFilePath);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    public void saveBytesToFile(byte[] bytes, String outputFilePath) {
        File file = new File(outputFilePath);
        File parentDir = file.getParentFile();

        // 检查目录是否存在，如果不存在则创建它
        if (parentDir != null && !parentDir.exists()) {
            if (!parentDir.mkdirs()) {
                try {
                    throw new IOException("无法创建目录: " + parentDir);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            fos.write(bytes);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


}