from cv2 import imread
import jieba
from wordcloud import WordCloud, ImageColorGenerator
import os

class DateAnalysis():
    def GetDate(num,bg_pic,save_path,font_path):
        global Date

        if num == 0:
            with open("data.main","r",encoding="utf-8") as f:
                Date = f.read()
                f.close()

            DateAnalysis.Analysis(Date,bg_pic,save_path,font_path)

            
        if num == 1:
            with open("date.read","r",encoding="utf-8") as f:
                Date = f.read()
                f.close()

            DateAnalysis.Analysis(Date,bg_pic,save_path,font_path)


    def get_current_txt(Date):
        sentences = ''

        for line in Date: # 数据清洗
            line = line.replace("\n", "") # 替换分行
            line = line.replace(" ", "") # 替换空格
            line = line.replace("\t", "") # 替换制表符
            line = line.replace("，", "") # 替换中文逗号
            line = line.replace("。", "") # 替换中文句号
            line = line.replace("？", "") # 替换中文问号
            line = line.replace("！", "") # 替换中文感叹号
            line = line.replace("：", "") # 替换中文冒号
            line = line.replace("（", "") # 替换中文左括号
            line = line.replace("）", "") # 替换中文右括号
            line = line.replace("《", "") # 替换中文左尖括号
            line = line.replace("》", "") # 替换中文右尖括号
            line = line.replace("“", "") # 替换中文双引号
            line = line.replace("”", "") # 替换中文双引号

            sentences = sentences + line # 拼接

        return sentences

    def Analysis(Date,bg_pic,save_path,font_path):
        Date = DateAnalysis.get_current_txt(Date)

        jieba.setLogLevel(jieba.logging.INFO) # 设置日志级别

        words = jieba.cut(Date,cut_all=True,HMM=True)
        
        # 统计词频
        word_dict = {}
        for word in words:
            if word in word_dict:
                word_dict[word] += 1
            else:
                word_dict[word] = 1
        
        # 对词频进行排序
        sort_list = sorted(word_dict.items(),key=lambda x:x[1],reverse=True)

        DateAnalysis.WordCloud(sort_list,word_dict,bg_pic,save_path,font_path)

    def WordCloud(sort_list,word_dict,bg_pic,save_path,font_path):
        # 制作词云
        word_list = []
        for i in sort_list:
            word_list.append(i[0])

        # 词云背景图片
        if bg_pic == "":
            bg_pic = 'bg.png'

        if save_path == "":
            save_path = os.path.join(os.path.expanduser('~'),"Desktop") + "/wordcloud.jpg"

        if font_path == "":
            font_path = 'msyh.ttc'

        bg = imread(bg_pic)

        wc = WordCloud(background_color="white",mask=bg,font_path=font_path,max_words=2000,max_font_size=100,random_state=42)
        wc.generate_from_frequencies(word_dict) 
        wc.to_file(save_path)

        with open(save_path + "data.txt","w",encoding="utf-8") as f:
            f.write(str(word_dict) + "\n") 
            f.close()
        