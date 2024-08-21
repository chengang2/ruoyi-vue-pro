package cn.iocoder.yudao.module.cg.framework.minio.config;

import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration(proxyBeanMethods = false)
@EnableConfigurationProperties(MinioProperties.class)
public class MinioConfiguration {

}
