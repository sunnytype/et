package com.sunnymix.et.excel.controller;

import com.sunnymix.et.excel.model.Book;
import com.sunnymix.et.excel.model.view.TextView;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.Optional;


@Controller
public class ExcelDisplayController {

    private final String uploadDir = "/tmp/upload/";

    @GetMapping("/excel/display")
    public String display(final Model model,
                          @RequestParam("textmode") Optional<Boolean> textmode,
                          @RequestParam("file") Optional<String> file) {

        file.ifPresent(fileName -> {
            File destFile = new File(uploadDir + fileName);

            if (destFile.exists() && destFile.isFile()) {
                Book book = Book.of(destFile);
                TextView textView = TextView.of(book);
                model.addAttribute("book", book);
                model.addAttribute("text", textView.getText());
            }
        });

        model.addAttribute("view", file.orElse(""));
        model.addAttribute("textmode", textmode.orElse(false));

        return "excel/display";
    }

    @PostMapping("/excel/display")
    public String displayPost(final Model model,
                              @RequestParam("textmode") Optional<Boolean> textmode,
                              @RequestParam("view") Optional<String> view,
                              @RequestParam("file") MultipartFile uploadFile) {
        String viewFile = "";

        String fileName = uploadFile.getOriginalFilename();
        if (fileName != null && !fileName.isBlank()) {
            File destFile = new File("/tmp/upload/" + fileName);

            if (!destFile.getParentFile().exists()) {
                boolean uploadFileCreated = destFile.getParentFile().mkdirs();
            }

            try {
                uploadFile.transferTo(destFile);
            } catch (Throwable fileError) {
                System.out.println(fileError.getMessage());
            }

            viewFile = fileName;
        }

        if (viewFile.isBlank()) {
            viewFile = view.orElse("");
        }

        boolean isTextmode = textmode.orElse(false);
        String textmodeParam = isTextmode ? "textmode=true" : "";

        return String.format("redirect:/excel/display?file=%s&%s", viewFile, textmodeParam);
    }

}
