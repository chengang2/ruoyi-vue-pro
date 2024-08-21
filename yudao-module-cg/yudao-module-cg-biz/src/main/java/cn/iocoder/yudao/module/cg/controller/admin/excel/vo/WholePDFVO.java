package cn.iocoder.yudao.module.cg.controller.admin.excel.vo;

import com.fasterxml.jackson.annotation.JsonProperty;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

@Schema(description = "管理后台 - cg创建最终pdf的VO")
@Data
public class WholePDFVO {

    @JsonProperty(value = "PDF_FM_SINGLE_FTP")
    @Schema(description = "PDF_FM_SINGLE_FTP")
    private String PDF_FM_SINGLE_FTP;
    @JsonProperty(value = "MERGE_PATH")
    @Schema(description = "MERGE_PATH")
    private String MERGE_PATH;
    @JsonProperty(value = "PDF_PRINT_PATH")
    @Schema(description = "PDF_PRINT_PATH")
    private String PDF_PRINT_PATH;
    @JsonProperty(value = "PDF_WATERMARK_PATH")
    @Schema(description = "PDF_WATERMARK_PATH")
    private String PDF_WATERMARK_PATH;


}
