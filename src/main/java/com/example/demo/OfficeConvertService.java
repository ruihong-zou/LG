package com.example.demo;

import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.document.DocumentFormat;
import org.jodconverter.core.document.DocumentFormatRegistry;
import org.springframework.stereotype.Service;

import java.io.*;

@Service
public class OfficeConvertService {

    private final DocumentConverter converter;
    private final DocumentFormatRegistry registry = DefaultDocumentFormatRegistry.getInstance();
    private final DocumentFormat DOC  = registry.getFormatByExtension("doc");
    private final DocumentFormat DOCX = registry.getFormatByExtension("docx");

    public OfficeConvertService(DocumentConverter converter) {
        this.converter = converter;
    }

    /** .doc (bytes) -> .docx (bytes) */
    public byte[] docToDocx(byte[] docBytes) throws Exception {
        try (InputStream in = new ByteArrayInputStream(docBytes);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            converter.convert(in).as(DOC).to(out).as(DOCX).execute();
            return out.toByteArray();
        }
    }

    /** .docx (bytes) -> .doc (bytes) */
    public byte[] docxToDoc(byte[] docxBytes) throws Exception {
        try (InputStream in = new ByteArrayInputStream(docxBytes);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            converter.convert(in).as(DOCX).to(out).as(DOC).execute();
            return out.toByteArray();
        }
    }
}
