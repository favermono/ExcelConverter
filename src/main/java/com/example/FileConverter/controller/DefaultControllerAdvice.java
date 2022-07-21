package com.example.FileConverter.controller;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.http.converter.HttpMessageNotReadableException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;

import java.io.IOException;

@ControllerAdvice
public class DefaultControllerAdvice {

    @ExceptionHandler({RuntimeException.class, HttpMessageNotReadableException.class})
    public ResponseEntity<String> exceptionHandler(RuntimeException e) {
        if (e.getClass() != HttpMessageNotReadableException.class){
            return new ResponseEntity<>(e.getMessage(), HttpStatus.BAD_REQUEST);
        } else {
            return new ResponseEntity<>("'url' field missed or some fields are filled in incorrectly (check README.md for more information).",
                    HttpStatus.BAD_REQUEST);
        }

    }
    @ExceptionHandler(IOException.class)
    public ResponseEntity<String> exceptionHandler(IOException e) {

        return new ResponseEntity<>("An error occurred while processing the file on server.",
                HttpStatus.INTERNAL_SERVER_ERROR); //Если ошибка возникла при работе с чтением или записью локальной копии файла
    }

}