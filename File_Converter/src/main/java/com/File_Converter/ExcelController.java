package com.File_Converter;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelProcessorService excelProcessorService;

    @GetMapping("/")
    public String welcome() {
        return "Welcome AWS";
    }

    
    public String handleFileUpload(@RequestParam("file") MultipartFile file,
            @RequestParam("outputDirectory") String outputDirectory) {
        try {
            excelProcessorService.processExcelFile(file.getOriginalFilename(), outputDirectory);
            return "File processed successfully.";
        } catch (IOException e) {
            e.printStackTrace();
            return "Error processing the file: " + e.getMessage();
        }
    }

    @GetMapping("/download/{fileName}")
    public ResponseEntity<FileSystemResource> downloadFile(@PathVariable String fileName) {
        String filePath = "C:\\Users\\Dell\\Downloads\\apollo\\" + fileName; 
        FileSystemResource resource = new FileSystemResource(filePath);

        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName);

        try {
			return ResponseEntity
			        .ok()
			        .headers(headers)
			        .contentLength(resource.contentLength())
			        .contentType(MediaType.APPLICATION_OCTET_STREAM)
			        .body(resource);
		} catch (IOException e) {
	
			e.printStackTrace();
		}
		return null;
    }
}
