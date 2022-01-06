import scrapy
from bs4 import BeautifulSoup


class NDHSpider(scrapy.Spider):
    name = 'ndh'
    start_urls = [
        'https://ndh.vn/doanh-nghiep'
    ]

    def parse_content(self, response):
        item = response.meta['item']
        soup = BeautifulSoup(response.body, 'html.parser')
        bodyElem = soup.select_one('div.list-news')
        item['content'] = ''
        if bodyElem is not None:
            item['content'] = bodyElem.get_text()
        item['domain'] = self.start_urls[0]
        return item

    def parse(self, response):
        for quote in response.css('h3.title-news > a'):
            detail_url = quote.css('a::attr("href")').extract_first()
            if detail_url is not None:
                item = {
                    'title': quote.css('a::text').extract_first(),
                    'url': detail_url,
                }
                request = scrapy.Request(response.urljoin(detail_url), callback=self.parse_content)
                request.meta['item'] = item
                yield request








