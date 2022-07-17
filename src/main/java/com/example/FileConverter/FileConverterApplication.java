package com.example.FileConverter;

import org.modelmapper.ModelMapper;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;


@SpringBootApplication
public class FileConverterApplication
{
	@Bean
	public ModelMapper modelMapper() {
		return new ModelMapper();
}

	public static void main(String[] args) throws Exception {
		SpringApplication.run(FileConverterApplication.class, args);
	}
}
