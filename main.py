# -*— codeing = utf-8 -*-

from bs4 import BeautifulSoup
import re
import urllib.request
import xlwt

findtitle = re.compile(r'<span class="title">(.*?)</span>')
findlink = re.compile(r'<a href="(.*?)">')
findpiclink = re.compile(r'<img alt=".*" class="" src="(.*?)" width="100"/>', re.S)
findbd = re.compile(r'<p class="">(.*?)</p>', re.S)
findrating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
findnum = re.compile(r'<span>(\d*)人评价</span>')
findinq = re.compile(r'<span class="inq">(.*?)</span>')


def main():
    baseurl = "https://movie.douban.com/top250?start="
    datalist = getdata(baseurl)
    savepath = "豆瓣电影Top250.xls"
    savedata(savepath, datalist)


def getdata(baseurl):
    datalist = []
    for i in range(10):
        i = i * 25
        i = str(i)
        url = baseurl + i
        askurl = Askurl(url)
        soup = BeautifulSoup(askurl, "html.parser")
        for message in soup.find_all('div', class_='item'):
            data = []
            message = str(message)

            title = re.findall(findtitle, message)
            if len(title) == 2:
                data.append(title[0])
                data.append(title[1])
            else:
                data.append(title[0])
                data.append(' ')

            link = re.findall(findlink, message)[0]
            data.append(link)

            pic_link = re.findall(findpiclink, message)[0]
            data.append(pic_link)

            rating = re.findall(findrating, message)[0]
            data.append(rating)

            num = re.findall(findnum, message)[0]
            data.append(num)

            inq = re.findall(findinq, message)
            if inq != 0:
                data.append(inq)
            else:
                data.append(' ')

            bd = re.findall(findbd, message)[0]
            bd = re.sub('<br/>', ' ', bd)
            data.append(bd)
            datalist.append(data)
    return datalist


def Askurl(url):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/88.0.4324.190 Safari/537.36 "
    }
    req = urllib.request.Request(url, headers=head)
    response = urllib.request.urlopen(req)
    urldata = response.read().decode('utf-8')
    return urldata


def savedata(savepath, datalist):
    workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = workbook.add_sheet('豆瓣TOP250', cell_overwrite_ok=True)

    order = ('电影名','其它名','电影链接','电影图片链接','评分','评分人数','概括','详细信息')
    for i in range(0,8):
        sheet.write(0,i,order[i])

    for i in range(0,250):
        for j in range(0,8):
            sheet.write(i+1,j,datalist[i][j])

    workbook.save(savepath)

if __name__ == "__main__":
    main()
