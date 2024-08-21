package cn.iocoder.yudao.module.cg.controller.app;

import cn.iocoder.yudao.framework.common.pojo.CommonResult;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import static cn.iocoder.yudao.framework.common.pojo.CommonResult.success;

@Tag(name = "用户 App - Test")
@RestController
@RequestMapping("/cg/test")
@Validated
public class AppDemoTestController {
    @GetMapping("/get")
    @Operation(summary = "获取 test 信息")
    public CommonResult<String> get() {
        System.out.println("cgtest1111........");
        return success("true");
    }

    //创建一个post接口，用于新增用户
    @PostMapping("/post")
    @Operation(summary = "获取 test 信息2")
    public CommonResult<String> post() {
        System.out.println("cgtest2222........");
        return success("true");
    }
}
