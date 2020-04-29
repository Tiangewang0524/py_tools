import requests
from lxml import etree
import re


headers = {
    'User-Agent':
        'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0;'
}

fp = open('C:/Users/administrator/Desktop/tpp.html', 'rb')

def tpp_url_list():
    tpp_url = []
    city_list = []

    r2 = fp.read().decode('utf-8')
    html = etree.HTML(r2)
    part_url = '//div[@class="M-cityList scrollStyle"]/ul'
    temp_url = html.xpath(part_url + '/li')

    for i in range(1, len(temp_url)):
        temp_li = "/li[" + str(i) + "]"
        temp_url_2 = html.xpath('//div[@class="M-cityList scrollStyle"]/ul/li[' + str(i) + ']/a')
        for j in range(1, len(temp_url_2)):
            temp_a = "/a[" + str(j) + "]/"
            combin_xpath = part_url + temp_li + temp_a
            city_name = html.xpath(combin_xpath + 'text()')
            city_url = html.xpath(combin_xpath + '/@href')
            tpp_url.extend(city_url)
            city_name = ''.join(city_name)
            city_list.append(city_name)
            city_url = ''.join(city_url)

    # print(city_list)
    # print(tpp_url)

    return tpp_url, city_list

def get_tpp_info(tpp_url, city_list):
    dict_city = dict()

    for i, url in enumerate(tpp_url):
        print('获取第{}个城市......'.format(i+1))
        r = requests.get(url, headers=headers)
        r.encoding = 'UTF-8'
        html = etree.HTML(r.text)
        try:
            # 判断是否有影片上映
            # movie_judge = html.xpath('//div[@class="tab-content"]/div[@class="tab-movie-list"]/div[@class="movie-card-wrap"]')
            # error_judge = html.xpath('//div[@class="tab-content"]/div[@class="tab-movie-list"]/div[@class="error-wrap"]')
            movie_part_xpath = '/html/body/div[4]/div[1]/div[2]/div[1]/div[@class="movie-card-wrap"]'
            movie_judge = html.xpath(movie_part_xpath)
            error_judge = html.xpath(
                '/html/body/div[4]/div[1]/div[2]/div[1]/div[@class="error-wrap"]')
            if movie_judge:
                temp_str = '有影片上映'
                movie_info = html.xpath(movie_part_xpath + '/a[1]/div[3]/span[1]/text()')
                for movie in movie_info:
                    temp_str += '影片名：' + str(movie)
                # print(movie_info)
                dict_city[city_list[i]] = temp_str
            if error_judge:
                temp_str = '无影片上映'
                dict_city[city_list[i]] = temp_str

            # print(dict_city)
        except:
            pass
    return dict_city


def main():
    tpp_url, city_list = tpp_url_list()
    tpp_info = get_tpp_info(tpp_url, city_list)
    tpp_order = sorted(tpp_info.items(), key=lambda x: x[1], reverse=False)
    print(tpp_order)


if __name__ == '__main__':
    main()
    # tpp_url_list()
