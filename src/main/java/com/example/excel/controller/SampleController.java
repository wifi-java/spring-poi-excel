package com.example.excel.controller;

import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
@RequiredArgsConstructor
public class SampleController {

  @GetMapping("/")
  public String sample() {
    return "sample";
  }
}
