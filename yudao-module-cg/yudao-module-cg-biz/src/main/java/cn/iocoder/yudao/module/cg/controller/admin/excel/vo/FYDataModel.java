package cn.iocoder.yudao.module.cg.controller.admin.excel.vo;

import com.fasterxml.jackson.annotation.JsonProperty;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

import java.util.List;
import java.util.Map;

@Schema(description = "管理后台 - cg创建附页VO")
@Data
public class FYDataModel {
    @JsonProperty(value = "FYStaticData")
    @Schema(description = "FYStaticData")
    private Map<String, String> FYStaticData;
    @JsonProperty(value = "FYDataTable")
    @Schema(description = "FYDataTable")
    private List<Map<String, String>> FYDataTable;
    @JsonProperty(value = "FYImageArray")
    @Schema(description = "FYImageArray")
    private List<String> FYImageArray;
    @JsonProperty(value = "FYColumnListToMerge")
    @Schema(description = "FYColumnListToMerge")
    private List<String> FYColumnListToMerge;
}
