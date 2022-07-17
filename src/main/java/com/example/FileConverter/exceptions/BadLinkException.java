package com.example.FileConverter.exceptions;

public class BadLinkException extends RuntimeException{
    public BadLinkException(String message)
    {
        super(message);
    }
}
