package cn.iocoder.yudao.module.cg.controller.admin.excel.vo;

import com.fasterxml.jackson.annotation.JsonProperty;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

import java.util.Map;

@Schema(description = "管理后台 - cg创建excelVO")
@Data
public class CreateExcel {
    @JsonProperty(value = "FOLDERNO")
    @Schema(description = "FOLDERNO")
    private String FOLDERNO;
    @JsonProperty(value = "PREORDNO")
    @Schema(description = "PREORDNO")
    private String PREORDNO;
    @JsonProperty(value = "REPORTGENTYPE")
    @Schema(description = "REPORTGENTYPE")
    private String REPORTGENTYPE;
    @JsonProperty(value = "INSPECTIONDATE")
    @Schema(description = "INSPECTIONDATE")
    private String INSPECTIONDATE;
    @JsonProperty(value = "QUALIFICATIONSFORDEPARTMENT")
    @Schema(description = "QUALIFICATIONSFORDEPARTMENT")
    private String QUALIFICATIONSFORDEPARTMENT;
    @JsonProperty(value = "MODEL_CODE_FMTD_PATH")
    @Schema(description = "MODEL_CODE_FMTD_PATH")
    private String MODEL_CODE_FMTD_PATH;
    @JsonProperty(value = "MODEL_CODE_FM_PATH")
    @Schema(description = "MODEL_CODE_FM_PATH")
    private String MODEL_CODE_FM_PATH;
    @JsonProperty(value = "MODEL_CODE_SY_PATH")
    @Schema(description = "MODEL_CODE_SY_PATH")
    private String MODEL_CODE_SY_PATH;
    @JsonProperty(value = "MODEL_CODE_FY_PATH")
    @Schema(description = "MODEL_CODE_FY_PATH")
    private String MODEL_CODE_FY_PATH;
    @JsonProperty(value = "MODEL_CODE2_PATH")
    @Schema(description = "MODEL_CODE2_PATH")
    private String MODEL_CODE2_PATH;
    @JsonProperty(value = "MODEL_CODE3_PATH")
    @Schema(description = "MODEL_CODE3_PATH")
    private String MODEL_CODE3_PATH;
    @JsonProperty(value = "EXCEL_FY_FTP")
    @Schema(description = "EXCEL_FY_FTP")
    private String EXCEL_FY_FTP;
    @JsonProperty(value = "Attached_PDF_Name")
    @Schema(description = "Attached_PDF_Name")
    private String Attached_PDF_Name;
    @JsonProperty(value = "FMTDDataList")
    @Schema(description = "FMTDDataList")
    private Map<String, String> FMTDDataList;
    @JsonProperty(value = "FMDataList")
    @Schema(description = "FMDataList")
    private Map<String, String> FMDataList;
    @JsonProperty(value = "SYDataList")
    @Schema(description = "SYDataList")
    private Map<String, String> SYDataList;
    @JsonProperty(value = "FYDataModel")
    @Schema(description = "FYDataModel")
    private FYDataModel FYDataModel;













}
