package com.example.FileConverter.odt;

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

    private String desired_sheets;
    private String row_to_skip;
    private String columns_to_skip;   // Номера столбцов через ","
    private boolean format_values;

    private String csv_format; // "CUSTOM" если поле не получено и "DEFAULT" если полученно значение, не входящее в данный список

    /** Эти параметры используются только если CSV_FORMAT = "CUSTOM" */
    private String value_separator;
    private Boolean first_line_is_header;
    private String quote_char;
    private String escape_char;
    private String comment_maker; //The comment start and the escape character cannot be the same ('1')
    private String null_string;
    private Boolean trim_fields;
    private String quote_mode;  // "MINIMAL" если поле не получено или некорректное значение
    private String record_separator;
    private Boolean trailing_delimiter;
    private Boolean allow_duplicate_header_names;

}
