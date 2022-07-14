package com.example.FileConverter.exceptions;

public class CorruptedFileException extends Exception{
    public CorruptedFileException(String message)
    {
        super(message);
    }
}
