package com.example.FileConverter.exceptions;

public class WrongFileFormatException extends Exception {
    public WrongFileFormatException(String message)
    {
        super(message);
    }
}