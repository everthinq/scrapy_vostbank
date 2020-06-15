from scrapy.cmdline import execute

def main():
    # execute(['scrapy', 'crawl', 'vostbank', '--nolog'])
    execute(['scrapy', 'crawl', 'vostbank'])

if __name__ == '__main__':
    main()