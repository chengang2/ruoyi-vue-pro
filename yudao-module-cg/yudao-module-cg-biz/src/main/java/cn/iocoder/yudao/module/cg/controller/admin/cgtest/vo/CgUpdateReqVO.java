package cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo;


import com.mzt.logapi.starter.annotation.DiffLogField;
import io.swagger.v3.oas.annotations.media.Schema;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Pattern;
import jakarta.validation.constraints.Size;
import lombok.Data;

@Schema(description = "管理后台 - cg创建/修改 Request VO")
@Data
public class CgUpdateReqVO {
    @Schema(description = "用户编号", example = "1024")
    private Long id;

    @Schema(description = "用户账号", requiredMode = Schema.RequiredMode.REQUIRED, example = "lims")
    @NotBlank(message = "用户账号不能为空")
    @Pattern(regexp = "^[a-zA-Z0-9]{4,30}$", message = "用户账号由 数字、字母 组成")
    @Size(min = 4, max = 30, message = "用户账号长度为 4-30 个字符")
    @DiffLogField(name = "用户账号")
    private String name;

    @Schema(description = "用户年龄", requiredMode = Schema.RequiredMode.REQUIRED, example = "30")
    @DiffLogField(name = "用户年龄")
    private Integer ageNumber;
}
