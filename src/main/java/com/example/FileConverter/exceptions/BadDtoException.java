package com.example.FileConverter.exceptions;

public class BadDtoException extends RuntimeException{
    public BadDtoException(String message)
    {
        super(message);
    }
}
