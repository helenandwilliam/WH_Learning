# -*- coding: utf-8 -*-
import scrapy

from rrs.items import RrsItem, FoodItem

class MeituanSpider(scrapy.Spider):
    name = "meituan"
    allowed_domains = ["cd.meituan.com"]
    start_urls = [ 
        'http://cd.meituan.com/category/meishi/'
    ]

    def parse(self, response):

        for sel in response.css("div.basic.cf > a"):
            rrs_item = RrsItem()
            rrs_item['r_name'] = sel.xpath('text()').extract()
            rrs_item['r_link'] = sel.xpath('@href').extract()
            #url = response.urljoin(sel.xpath('@href').extract())

            yield scrapy.Request(rrs_item['r_link'][0], meta = {'rrs_item' : rrs_item},
                    callback = self.parse_food)

        next_page = response.css("li.next > a::attr('href')")

        if next_page:
            url = response.urljoin(next_page[0].extract())
            yield scrapy.Request(url, self.parse)



    def parse_food(self, response):

        rrs_item = response.meta['rrs_item']

        r_foods = []

        for sel in response.css("div.menu__items > table > tr > td"):
            r_food = FoodItem()
            r_food['f_name'] = sel.xpath('text()').extract()
            r_food['f_price'] = sel.xpath('span/text()').extract()

            r_foods.append(r_food)

        rrs_item['r_foods'] = r_foods

        yield rrs_item
