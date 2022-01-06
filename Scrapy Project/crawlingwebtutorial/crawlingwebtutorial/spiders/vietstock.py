import scrapy
from bs4 import BeautifulSoup


class VietStockSpider(scrapy.Spider):
    name = 'vietstock'
    start_urls = [
        'https://vietstock.vn/doanh-nghiep.htm'
    ]

    def parse_content(self, response):
        item = response.meta['item']
        soup = BeautifulSoup(response.body, 'html.parser')
        bodyElem = soup.select_one('div.channel-container')
        item['content'] = ''
        if bodyElem is not None:
            item['content'] = bodyElem.get_text()
        item['domain'] = self.start_urls[0]
        return item

    def parse(self, response):
        for quote in response.css('h2.channel-title > a'):
            detail_url = quote.css('a::attr("href")').extract_first()
            if detail_url is not None:
                item = {
                    'title': quote.css('a::text').extract_first(),
                    'url': detail_url,
                }
                request = scrapy.Request(response.urljoin(detail_url), callback=self.parse_content)
                request.meta['item'] = item
                yield request
            next_page = response.xpath('//*[@id="page-next "]/a').extract()
            if next_page:
                next_href = next_page[0]
                next_page_url = 'https://vietstock.vn/doanh-nghiep.htm' + next_href
                request = scrapy.Request(url=next_page_url)
                yield request








