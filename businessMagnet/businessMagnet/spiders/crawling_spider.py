from scrapy.spiders import CrawlSpider, Rule
from scrapy.linkextractors import LinkExtractor
import pandas as pd
import scrapy
from urllib.parse import urljoin
from openpyxl.workbook import Workbook

class CrawlingSpider(CrawlSpider):
    name = "crawler"
    allowed_domains = ["businessmagnet.co.uk"]
    #start_urls = ["https://www.businessmagnet.co.uk/company/a.htm"]
    scraped_data = []

    def start_requests(self):
        base_url = 'https://www.businessmagnet.co.uk/company/'
        letters = [chr(i) for i in range(ord('w'), ord('z') + 1)]

        for letter in letters:
            url = f'{base_url}{letter}.htm'
            yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'a':
                for i in range(2, 39):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'b':
                for i in range(2, 30):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'c':
                for i in range(2, 44):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'd':
                for i in range(2, 22):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'e':
                for i in range(2, 22):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'f':
                for i in range(2, 18):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'g':
                for i in range(2, 18):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'h':
                for i in range(2, 19):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'i':
                for i in range(2, 16):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'j':
                for i in range(2, 12):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'k':
                for i in range(2, 11):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'l':
                for i in range(2, 17):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'm':
                for i in range(2, 30):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'n':
                for i in range(2, 13):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'o':
                for i in range(2, 10):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'p':
                for i in range(2, 32):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'q':
                for i in range(2, 4):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'r':
                for i in range(2, 22):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 's':
                for i in range(2, 46):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 't':
                for i in range(2, 30):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'u':
                for i in range(2, 7):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'v':
                for i in range(2, 8):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'w':
                for i in range(2, 19):
                    url = f'{base_url}{letter}/page{i}'
                    yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'y':
                url = f'{base_url}{letter}/page2'
                yield scrapy.Request(url, callback=self.parse_company_list)
            if letter == 'z':
                url = f'{base_url}{letter}/page2'
                yield scrapy.Request(url, callback=self.parse_company_list)

    #rules = (
    #    Rule(LinkExtractor(allow="/company/[a-z]\.htm$", deny=()), follow=True),
    #    Rule(LinkExtractor(allow="/company/.*\.htm$", deny=()), callback='parse_company_page')
    #)

    def parse_company_list(self, response):
        company_links = response.css('section.indexinglist div.midcontainer ul li a::attr(href)').getall()
        for company_link in company_links:
            absolute_url = urljoin(response.url, company_link)
            yield scrapy.Request(absolute_url, callback=self.parse_company_page)

    def parse_company_page(self, response):
        phone_number = response.css('div.company-phno span::text').get()
        address_lines = response.css('p.text-xs::text').getall()
        address = ' '.join(line.strip() for line in address_lines if line.strip())
        website = response.css('span[class~="inline-block"][class~="text-lg"][class~="px-4"][class~="truncate"]::text').get()
        services = []
        a_elements = response.css('ol.max-h-\[240px\].overflow-y-auto.text-xs.nice-scrollbar li a::text').getall()

        result = ', '.join(a_elements)


        if phone_number:
            phone_number=phone_number.strip()


        data = {
            "name": response.css("h1::text").get().replace('\n', ''),
            "address": address,
            "website": website,
            "contact": phone_number,
            "services": result
        }
        self.scraped_data.append(data)

    def closed(self, reason):
        df = pd.DataFrame(self.scraped_data)
        df.to_excel('output.xlsx', index=False)
