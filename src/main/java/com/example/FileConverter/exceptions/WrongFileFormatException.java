package com.example.FileConverter.exceptions;

public class WrongFileFormatException extends RuntimeException {
    public WrongFileFormatException(String message)
    {
        super(message);
    }
}
