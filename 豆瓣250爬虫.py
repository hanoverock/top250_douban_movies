import requests,openpyxl,time
from bs4 import BeautifulSoup


wb=openpyxl.Workbook()
sheet=wb.active
sheet.title='豆瓣250'
sheet['A1']='序号'
sheet['B1']='片名'
sheet['C1']='评分'
sheet['D1']='推荐语'
sheet['E1']='链接'

cell_list=[]
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}

for i in range(0,10):
    res=requests.get('https://movie.douban.com/top250?start=%d&filter='%(i*25),headers=headers)
    html=res.text
    soup=BeautifulSoup(html,'html.parser')

    body=soup.find('ol')
    units=body.find_all('li')
    for item in units:
        #序号
        number=item.find('em').text
        #片名
        nametaglist=item.find_all(class_='title')#各种译名的tag
        name_list=[]#创建空表，方便取出片名后添加添加
        for tag in nametaglist:
            name=tag.text#循环找单个译名
            name_list.append(name)#译名添加至空表
        names=' '.join(name_list)#对添加完的列表用jion函数，以便在单元格内并列
        #评分
        rating=str(item.find(class_='rating_num').text)+'分'
        #推荐语
        comment_tag=item.find(class_='inq')
        if comment_tag==None:#应对没有推荐语的情况
            comment='暂无'
        else:
            comment=comment_tag.text
        #链接
        url=item.find('a')['href']
        #装进列表，以便写入
        row_list=[number,names,rating,comment,url]
        sheet.append(row_list)
        #cell_list.append(row_list)#也可以存放在大列表，然后在循环结束后写入

time=time.asctime(time.localtime(time.time()) )
sheet.append([time])
wb.save('豆瓣250.xlsx')

    




