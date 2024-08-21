package cn.iocoder.yudao.module.cg.framework.web.config;

import cn.iocoder.yudao.framework.swagger.config.YudaoSwaggerAutoConfiguration;
import org.springdoc.core.models.GroupedOpenApi;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration(proxyBeanMethods = false)
public class CgWebConfiguration {

    /**
     * cg 模块的 API 分组
     */
    @Bean
    public GroupedOpenApi cgGroupedOpenApi() {
        return YudaoSwaggerAutoConfiguration.buildGroupedOpenApi("cg");
    }

}
