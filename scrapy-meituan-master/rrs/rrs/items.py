# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

from scrapy.item import Item, Field


class RrsItem(Item):
    r_name = Field()
    r_link = Field()
    r_foods = Field()


class FoodItem(Item):
    f_name = Field()
    f_price = Field()
