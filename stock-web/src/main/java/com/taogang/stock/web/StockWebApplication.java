package com.taogang.stock.web;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class StockWebApplication {

	public static void main(String[] args) {
		Logger log = LoggerFactory.getLogger(StockWebApplication.class);
		log.info("web 服务启动开始");
		SpringApplication.run(StockWebApplication.class, args);
		log.info("web 服务启动结束");
	}

}
