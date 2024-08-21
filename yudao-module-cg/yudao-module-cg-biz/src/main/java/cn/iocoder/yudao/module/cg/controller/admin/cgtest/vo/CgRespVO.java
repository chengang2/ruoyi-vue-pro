package cn.iocoder.yudao.module.cg.controller.admin.cgtest.vo;


import com.alibaba.excel.annotation.ExcelProperty;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

import java.time.LocalDateTime;

@Schema(description = "管理后台 - cg创建/修改 Request VO")
@Data
public class CgRespVO {
    @Schema(description = "用户编号", example = "1024")
    @ExcelProperty("用户编号")
    private Long id;
    @Schema(description = "用户账号", requiredMode = Schema.RequiredMode.REQUIRED, example = "lims")
    @ExcelProperty("用户账号")
    private String name;
    @Schema(description = "用户年龄", requiredMode = Schema.RequiredMode.REQUIRED, example = "30")
    @ExcelProperty("用户年龄")
    private Integer ageNumber;
    /**
     * 创建时间
     */
    @Schema(description = "创建时间", example = "yyyy-MM-dd HH:mm:ss")
    @JsonSerialize(using = CustomLocalDateTimeSerializer.class)
    private LocalDateTime createTime;
    /**
     * 最后更新时间
     */
    @Schema(description = "更新时间", example = "yyyy-MM-dd HH:mm:ss")
    private LocalDateTime updateTime;
    /**
     * 创建者，目前使用 SysUser 的 id 编号
     *
     * 使用 String 类型的原因是，未来可能会存在非数值的情况，留好拓展性。
     */
    @Schema(description = "创建者id", example = "1")
    private String creator;
    @Schema(description = "创建者名称", example = "张三")
    private String creatorName;
    /**
     * 更新者，目前使用 SysUser 的 id 编号
     *
     * 使用 String 类型的原因是，未来可能会存在非数值的情况，留好拓展性。
     */
    @Schema(description = "修改者id", example = "1")
    private String updater;
    @Schema(description = "修改者名称", example = "张三")
    private String updaterName;
    /**
     * 多租户编号
     */
    @Schema(description = "租户编号", example = "2")
    private Long tenantId;
}
