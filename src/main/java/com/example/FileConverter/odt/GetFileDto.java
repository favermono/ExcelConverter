package com.example.FileConverter.odt;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.springframework.web.multipart.MultipartFile;

import io.swagger.v3.oas.annotations.media.Schema;
import lombok.Builder;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;


@Getter
@Setter
@Builder
public class GetFileDto {

    public enum MyEnum {
        CUSTOM, TWO;

    }

    @NonNull
    private String URL;

    private String DESIRED_SHEETS;
    @Builder.Default
    private int ROWS_TO_SKIP=0;
    private String COLUMNS_TO_SKIP;   // Номера столбцов через ","
    @Builder.Default
    private boolean FORMAT_VALUES=false;

    @Builder.Default
    @Schema(type = "string", allowableValues = { "CUSTOM", "RFC_4180", "EXCEL", "TDF", "MYSQL", "INFORMIX_UNLOAD", "INFORMIX_UNLOAD_CSV" })
    private String CSV_FORMAT = "CUSTOM";


    /* Эти параметры используются только если CSV_FORMAT = "CUSTOM" */
    @Builder.Default
    private String VALUE_SEPARATOR=",";
    @Builder.Default
    private Boolean FIRST_LINE_IS_HEADER=true;
    @Builder.Default
    private String QUOTE_CHAR ="\"";
    @Builder.Default
    private String ESCAPE_CHAR="\\"; //sam
    private String COMMENT_MAKER;
    private String NULL_STRING;
    @Builder.Default
    private Boolean TRIM_FIELDS =true;
    @Builder.Default
    @Schema(type = "string", allowableValues = { "ALL", "MINIMAL", "NON_NUMERIC", "NONE" })
    private String QUOTE_MODE = "MINIMAL";
    @Builder.Default
    private String RECORD_SEPARATOR="\\n";
    @Builder.Default
    private Boolean TRAILING_DELIMITER=false;
    @Builder.Default
    private Boolean ALLOW_DUPLICATE_HEADER_NAMES=true;

}
