package com.example.FileConverter.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.*;


import lombok.Builder;
import lombok.Getter;
import lombok.Setter;


@Getter
@Setter
@Builder

public class GetFileDto {


    @NonNull
    private String url;
    @JsonProperty("desired_sheets")
    private String desiredSheets;
    @JsonProperty("rows_to_skip")
    private String rowsToSkip;
    @JsonProperty("columns_to_skip")
    private String columnsToSkip;   // Номера столбцов через ","
    @JsonProperty("format_values")
    private boolean formatValues;

    @JsonProperty("csv_format")
    private String csvFormat; // "CUSTOM" если поле не получено и "DEFAULT" если полученно значение, не входящее в данный список

    /** Эти параметры используются только если CSV_FORMAT = "CUSTOM" */
    @JsonProperty("value_separator")
    private String valueSeparator;
    @JsonProperty("first_line_is_header")
    private Boolean firstLineIsHeader;
    @JsonProperty("quote_char")
    private String quoteChar;
    @JsonProperty("escape_char")
    private String escapeChar;
    @JsonProperty("comment_maker")
    private String commentMaker; //The comment start and the escape character cannot be the same ('1')
    @JsonProperty("null_string")
    private String nullString;
    @JsonProperty("trim_fields")
    private Boolean trimFields;
    @JsonProperty("quote_mode")
    private String quoteMode;  // "MINIMAL" если поле не получено или некорректное значение
    @JsonProperty("record_separator")
    private String recordSeparator;
    @JsonProperty("trailing_delimiter")
    private Boolean trailingDelimiter;
    @JsonProperty("allow_duplicate_header_names")
    private Boolean allowDuplicateHeaderNames;

}
