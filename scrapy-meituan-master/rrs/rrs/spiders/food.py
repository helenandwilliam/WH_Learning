# -*- coding: utf-8 -*-
import scrapy

from rrs.items import RrsItem, FoodItem

class FoodSpider(scrapy.Spider):
    name = "food"
    allowed_domains = ["cd.meituan.com"]
    start_urls = [ 
        'http://cd.meituan.com/shop/42105046'
    ]

    def parse(self, response):

        for sel in response.css("div.menu__items > table > tr > td"):
            food = FoodItem()
            food['f_name'] = sel.xpath('text()').extract()
            food['f_price'] = sel.xpath('span/text()').extract()

            yield food

