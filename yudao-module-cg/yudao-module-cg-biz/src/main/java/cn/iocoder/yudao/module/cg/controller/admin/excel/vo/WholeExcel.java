package cn.iocoder.yudao.module.cg.controller.admin.excel.vo;

import com.fasterxml.jackson.annotation.JsonProperty;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

@Schema(description = "管理后台 - cg创建最终pdf的VO")
@Data
public class WholeExcel {

    @JsonProperty(value = "FOLDERNO")
    @Schema(description = "FOLDERNO")
    private String FOLDERNO;
    @JsonProperty(value = "PREORDNO")
    @Schema(description = "PREORDNO")
    private String PREORDNO;
    @JsonProperty(value = "INSPECTIONDATE")
    @Schema(description = "INSPECTIONDATE")
    private String INSPECTIONDATE;
    @JsonProperty(value = "QFDATE")
    @Schema(description = "QFDATE")
    private String QFDATE;
    @JsonProperty(value = "EXCEL_FM_SINGLE_FTP")
    @Schema(description = "EXCEL_FM_SINGLE_FTP")
    private String EXCEL_FM_SINGLE_FTP;
    @JsonProperty(value = "EXCEL_EDIT_PATH")
    @Schema(description = "EXCEL_EDIT_PATH")
    private String EXCEL_EDIT_PATH;
    @JsonProperty(value = "EXCEL_FM_FTP")
    @Schema(description = "EXCEL_FM_FTP")
    private String EXCEL_FM_FTP;
    @JsonProperty(value = "EXCEL_SY_FTP")
    @Schema(description = "EXCEL_SY_FTP")
    private String EXCEL_SY_FTP;
    @JsonProperty(value = "EXCEL_FY_FTP")
    @Schema(description = "EXCEL_FY_FTP")
    private String EXCEL_FY_FTP;
    @JsonProperty(value = "EXCEL_FY2_PATH")
    @Schema(description = "EXCEL_FY2_PATH")
    private String EXCEL_FY2_PATH;
    @JsonProperty(value = "EXCEL_FY3_PATH")
    @Schema(description = "EXCEL_FY3_PATH")
    private String EXCEL_FY3_PATH;
    @JsonProperty(value = "Attached_PDF_Name")
    @Schema(description = "Attached_PDF_Name")
    private String Attached_PDF_Name;
    @JsonProperty(value = "StampBite")
    @Schema(description = "StampBite")
    private String StampBite;
    @JsonProperty(value = "StampBite_QF")
    @Schema(description = "StampBite_QF")
    private String StampBiteQF;
    @JsonProperty(value = "sTestBy")
    @Schema(description = "sTestBy")
    private String sTestBy;
    @JsonProperty(value = "sApprovedBy")
    @Schema(description = "sApprovedBy")
    private String sApprovedBy;
    @JsonProperty(value = "sReleasedBy")
    @Schema(description = "sReleasedBy")
    private String sReleasedBy;

}
