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
    private String URL;

    private String DESIRED_SHEETS;
    private int ROWS_TO_SKIP;
    private String COLUMNS_TO_SKIP;   // Номера столбцов через ","
    private boolean FORMAT_VALUES;

    private String CSV_FORMAT; // "CUSTOM" если поле не получено и "DEFAULT" если полученно значение, не входящее в данный список

    /** Эти параметры используются только если CSV_FORMAT = "CUSTOM" */
    private String VALUE_SEPARATOR;
    private Boolean FIRST_LINE_IS_HEADER;
    private String QUOTE_CHAR;
    private String ESCAPE_CHAR;
    private String COMMENT_MAKER; //The comment start and the escape character cannot be the same ('1')
    private String NULL_STRING;
    private Boolean TRIM_FIELDS;
    private String QUOTE_MODE;  // "MINIMAL" если поле не получено или некорректное значение
    private String RECORD_SEPARATOR;
    private Boolean TRAILING_DELIMITER;
    private Boolean ALLOW_DUPLICATE_HEADER_NAMES;

}
