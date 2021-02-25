# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy
from scrapy.loader.processors import MapCompose, TakeFirst


def remove_quotes(text):
    text = text.strip(u'\u201c'u'\u201d')
    return text


class QuotesItem(scrapy.Item):
    # define the fields for your item here like:
    # name = scrapy.Field()
    author = scrapy.Field()
    quote = scrapy.Field(
        input_processor=MapCompose(remove_quotes),
        output_processor=TakeFirst()
    )
    tags = scrapy.Field()
