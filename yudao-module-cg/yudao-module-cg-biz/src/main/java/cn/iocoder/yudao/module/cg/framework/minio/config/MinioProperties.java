package cn.iocoder.yudao.module.cg.framework.minio.config;

import io.minio.MinioClient;
import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.validation.annotation.Validated;

@ConfigurationProperties(prefix = "yudao.minio")
@Validated
@Data
public class MinioProperties {

    private String accessKey;

    private String secretKey;

    private String endpoint;

    private String bucketName;

    private String reportTemp;

    @Bean
    public MinioClient minioClient() {
        return MinioClient.builder()
                .endpoint(endpoint)
                .credentials(accessKey, secretKey)
                .build();
    }
}
