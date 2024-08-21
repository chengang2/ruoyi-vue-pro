package cn.iocoder.yudao.module.cg.controller.admin.excel.vo;


import com.alibaba.excel.annotation.ExcelProperty;
import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Data;

@Schema(description = "管理后台 - cg创建返回VO")
@Data
public class CreateRespVO {

    @Schema(description = "封面套打xls", example = "封面套打.xls")
    @ExcelProperty("封面套打excel")
    private String fmtdExcelPath;

//    @Schema(description = "封面套打pdf", example = "封面套打.pdf")
//    @ExcelProperty("封面套打pdf")
//    private String fmtdPdfPath;

    @Schema(description = "封面xls", example = "封面.xls")
    @ExcelProperty("封面excel")
    private String fmExcelPath;

//    @Schema(description = "封面pdf", example = "封面.pdf")
//    @ExcelProperty("封面pdf")
//    private String fmPdfPath;

    @Schema(description = "首页xls", example = "首页.xls")
    @ExcelProperty("首页excel")
    private String syExcelPath;

//    @Schema(description = "首页pdf", example = "首页.pdf")
//    @ExcelProperty("首页pdf")
//    private String syPdfPath;

    @Schema(description = "附页xls", example = "附页.xls")
    @ExcelProperty("附页excel")
    private String fyExcelPath;

//    @Schema(description = "附页pdf", example = "附页.pdf")
//    @ExcelProperty("附页pdf")
//    private String fyPdfPath;

//    @Schema(description = "过程中的pdf", example = "xxxx.pdf")
//    @ExcelProperty("过程中的pdf")
//    private String path2;

}
