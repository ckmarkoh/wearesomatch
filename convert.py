# -*- coding: utf-8 -*
import xlrd
import dicttoxml
import re
from urlparse import urljoin

_MAPPING_={
0 : 'id',
1 : 'name',
2 : 'website',
3 : 'address',
4 : 'type',
5 : 'desc',
6 : 'other',

7 : 'job1type',
8 : 'job1title',
9 : 'job1duty',
10: 'job1cons',
11: 'job1inter',
12: 'job1hassal',
13: 'job1sal',

14: 'job2type',
15: 'job2title',
16: 'job2duty',
17: 'job2cons',
18: 'job2inter',
19: 'job2hassal',
20: 'job2sal',

21: 'job3type',
22: 'job3title',
23: 'job3duty',
24: 'job3cons',
25: 'job3inter',
26: 'job3hassal',
27: 'job3sal',

28: 'job4type',
29: 'job4title',
30: 'job4duty',
31: 'job4cons',
32: 'job4inter',
33: 'job4hassal',
34: 'job4sal',

35: 'job5type',
36: 'job5title',
37: 'job5duty',
38: 'job5cons',
39: 'job5inter',
40: 'job5hassal',
41: 'job5sal',

42: 'job6type',
43: 'job6title',
44: 'job6duty',
45: 'job6cons',
46: 'job6inter',
47: 'job6hassal',
48: 'job6sal',

49: 'job7type',
50: 'job7title',
51: 'job7duty',
52: 'job7cons',
53: 'job7inter',
54: 'job7hassal',
55: 'job7sal',

56: 'job8type',
57: 'job8title',
58: 'job8duty',
59: 'job8cons',
60: 'job8inter',
61: 'job8hassal',
62: 'job8sal',
}

_INTERN_DICT={ 
1.  :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/%E9%AB%98%E6%A0%A1%E8%AA%8C%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8D%E6%99%AF%E6%96%87%E7%A7%91%E5%A4%A7%E6%A5%8A%E7%AB%8B%E6%A5%B7/',' 高校誌實習心得－景文科大楊立楷  '),
14. :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/patisco%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8Dchayin/  ',' Patisco實習心得－Chayin '),
20. :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/%E5%93%81%E5%94%84%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8D%E6%94%BF%E5%A4%A7%E4%B8%AD%E6%96%87%E5%AE%8B%E9%A8%8F%E5%9D%87/',' 品唄實習心得－政大中文宋騏均    '),
23. :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/arpp%E7%90%85%E8%8C%B6%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8D%E5%8F%B0%E5%A4%A7%E7%B6%93%E6%BF%9F%E5%91%A8%E6%9B%89%E5%90%9F/',' ARPP(琅茶)實習心得－台大經濟周曉吟  '),
25. :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/%E6%97%8B%E8%BD%89%E6%9C%A8%E9%A6%AC%E5%AF%A6%E7%BF%92%E7%94%9F%EF%BC%8D%E5%90%B3%E6%99%89%E5%AE%87/',' 旋轉木馬實習生－吳晉宇  '),
32. :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/%E6%96%B9%E7%99%BD%E7%A7%91%E6%8A%80%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8D%E5%B0%8F%E6%8D%B2/',' 方白科技實習心得－小捲  '),
41. :(r'http://intern.grazingcat.com/intern-highlights/intern-experience.html   ',' Grazingcat Intern 牧貓實習  '),
43. :(r'http://ntueawearesomatch3.wordpress.com/2013/11/06/%E8%89%BE%E7%89%B9%E7%B6%B2%E8%B7%AF%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8D%E5%8C%97%E6%95%99%E5%A4%A7-%E7%8E%A9%E5%85%B7%E8%88%87%E9%81%8A%E6%88%B2%E8%A8%AD%E8%A8%88%E9%84%AD%E5%AF%B6%E7%90%B3/',' 艾特網路實習心得－北教大 鄭寶琳 '),
46.  :(r'http://ntueawearesomatch3.wordpress.com/2014/03/28/pobono%E5%AF%A6%E7%BF%92%E5%BF%83%E5%BE%97%EF%BC%8D%E9%BB%83%E6%9F%8F%E7%BF%B0/',' Pobono實習心得－黃柏翰  '),
}

_LOGO_DICT = {
1.:'1.jpg',
10.:'10.jpg',
13.:'13.png',
14.:'14.jpg',
15.:'15.jpg',
16.:'16.png',
18.:'18.jpg',
19.:'19.jpg',
20.:'20.jpg',
21.:'21.jpg',
23.:'23.jpg',
24.:'24.jpg',
26.:'26.png',
27.:'27.jpg',
30.:'30.jpg',
31.:'31.jpg',
32.:'32.jpg',
33.:'33.jpg',
34.:'34.png',
35.:'35.jpg',
36.:'36.jpg',
37.:'37.jpg',
39.:'39.jpg',
40.:'40.jpg',
41.:'41.png',
43.:'43.png',
44.:'44.jpg',
46.:'46.jpg',
5.:'5.jpg',
9.:'9.jpg',
}


_NDIS_LIST = [ 48.,49.,50.,51. ]


_NKEY='B533C1C863AE4C'
def read_data():
    data = xlrd.open_workbook("data.xlsx")
    #print data.sheet_names()[0]
    sh = data.sheet_by_index(0)
    return sh


def replace_multi(rstr, rlist, resub=False):
    rstr2=rstr
    for rl in rlist:
        if not resub:
            rstr2=rstr2.replace(rl[0],rl[1])
        else:
            rstr2=re.sub(rl[0],rl[1],rstr2) 
    return rstr2
        

def main():
    sh = read_data()
 #   i=0
 #   for x in sh.row(0):
 #       print i,x.value
 #       i+=1
    max_range=58
    temp_ary=[]
    #len_ary=[]
    for i in range(1,max_range+1):
        temp_dict={}
        for j in range(0,sh.ncols): 
            elem=sh.row(i)[j]
            if j%7==0 and j!=0:
                val=replace_multi(elem.value,[(r'\([^\(\)]*\)','')],True)
            elif j==2:
                val=urljoin('http:', elem.value).replace('///','//') 
            elif j==6:
                val=replace_multi(elem.value,[('\n',_NKEY)])
            else:
                val=elem.value
            #if val !="":
            temp_dict.update({_MAPPING_[j]:val})    
        this_id = float(temp_dict['id'])
        if this_id in _INTERN_DICT.keys():
            temp_dict.update({'intern_url':_INTERN_DICT[this_id][0]})    
            temp_dict.update({'intern_title':_INTERN_DICT[this_id][1].decode('utf-8')})    
        if this_id in _LOGO_DICT.keys():
            temp_dict.update({'logo':_LOGO_DICT[this_id]})    
        if this_id in _NDIS_LIST:
            temp_dict.update({'display':'否'.decode('utf-8')})
        else:
            temp_dict.update({'display':'是'.decode('utf-8')})
        temp_ary.append({'company':temp_dict})
    result_s=dicttoxml.dicttoxml(temp_ary).encode('utf-8')
    #print [result_s]
    result_s=replace_multi(result_s,[('<item type="dict">',''),('</item>',''),
                                     ('<root>','<companys>'),('</root>','</companys>'),(_NKEY,'<br/>')])
    result_s=replace_multi(result_s,[(r'type="[^"]*"','')],True)
    print result_s
    #print max(len_ary)
    
         





if __name__ == "__main__":
    main()
