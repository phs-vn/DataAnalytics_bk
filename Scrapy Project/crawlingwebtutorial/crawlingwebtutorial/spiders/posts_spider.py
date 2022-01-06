import scrapy


class PostsSpider(scrapy.Spider):
    name = "posts"

    start_urls = [
        'https://cafef.vn/',
    ]

    def parse(self, response, **kwargs):
        page = response.url.split('/')[-1]
        filename = f"posts-{page}.html"
        with open(filename, 'wb') as f:
            f.write(response.body)
