package com.sunnymix.et;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@SpringBootApplication
@Controller
public class EtApplication {

    public static void main(String[] args) {
        SpringApplication.run(EtApplication.class, args);
    }

    @GetMapping("/")
    public String index() {
        return "index";
    }

}
