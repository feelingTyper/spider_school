import urllib
import urllib.request
from bs4 import BeautifulSoup
from spider_class import Spider
import time

# 每一所大学有一个这样的类
class Rice(Spider):

    def __init__(self):
        Spider.__init__(self)

    def rice_spider(self, pinyin, pinyin_dict):
        '''
        爬取主页入口的文件，包含提取规则等等
        :param pinyin:
        :param pinyin_dict:
        :return:
        '''
        url = 'https://search.rice.edu/html/people/p/0/0/?lastname={}'.format(urllib.parse.quote(pinyin))
        print('url is %s' % url)
        student_list = []
        name_dict = pinyin_dict
        name_dict = list(map(lambda x: x.lower(), name_dict))
        try:
            req = urllib.request.Request(url)
            source_code = urllib.request.urlopen(req).read()
            plain_text = str(source_code)
            soup = BeautifulSoup(plain_text, "lxml")
            result_soup = soup.find('div', {'id': 'results'})

            if not result_soup:
                student_list.append(['-', '-', '-', 0])
                return student_list

            for res_item in result_soup.findAll('div', {'id': 'peopleresults'}):
                name = res_item.find('a', {'class': 'name'}).text
                # name = self.parse_content(name)
                peopleinfo = res_item.find('div', {'class': 'peopleinfo'})
                email = peopleinfo.find('p', {'class': 'email'})
                if not email:
                    continue

                email = email.a.text
                grade = peopleinfo.find('p', {'class': 'year'})
                if not grade:
                    continue

                grade = grade.text
                chinese = self.is_chinese(name, name_dict)
                print(chinese)
                student_list.append([name, grade, email, chinese])

            return student_list
        except (urllib.request.HTTPError, urllib.request.URLError) as e:
            print(e)
            return -1

    def spider(self):
        '''
        开始爬取函数
        :return:
        '''
        LOG_FILE = 'rice_class.log'
        logname = 'rice_class'
        # 获取日志
        logger = self.getlog(LOG_FILE, logname)
        # 统计从哪个位置开始接着爬
        count = self.get_count(LOG_FILE)
        save_path = 'rice/'
        output_path = 'output/rice/'
        pinyin_list = self.name_dict
        while count < len(pinyin_list):
            book_lists = -1
            print('now spidering pinyin is %s  && number is %d' % (pinyin_list[count], count))
            logger.info('now spidering pinyin is {}  && number is {}'.format(pinyin_list[count], count))
            logger.info(count)
            while book_lists == -1:
                book_lists = self.rice_spider(pinyin_list[count], self.name_dict)
            self.print_book_lists_excel(book_lists, pinyin_list[count], save_path)
            count += 1
            print('sleeping.......')
            time.sleep(5)
        # 合并所有xlsx文件
        if count == len(pinyin_list):
            print('merging...')
            self.compct_xlsx_py(save_path, output_path)
            self.compct_xlsx_all(save_path, output_path)
            self.compct_xlsx_all_chinese(save_path, output_path)
            print('merge finished')

