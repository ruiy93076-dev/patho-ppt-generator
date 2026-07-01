# patho-ppt-generator
from pptx import Presentation
from pptx.util import Inches, Pt

# 新建空白PPT
prs = Presentation()

# 自定义单页模板：疾病标题+考点文字+左大体图(网络链接)+右镜下切片(网络链接)
def add_pathology_slide(title, knowledge_text, macro_pic_url, slice_pic_url):
    # 选用双栏布局
    slide_layout = prs.slide_layouts[3]
    slide = prs.slides.add_slide(slide_layout)

    # 设置页面标题
    slide.shapes.title.text = title
    title_font = slide.shapes.title.text_frame.paragraphs[0].font
    title_font.size = Pt(24)
    title_font.bold = True

    # 左侧填入背诵考点
    text_box = slide.placeholders[1]
    text_box.text = knowledge_text
    for p in text_box.text_frame.paragraphs:
        p.font.size = Pt(14)

    # 插入左侧：大体标本图（直接读取网络链接，无需上传本地图片）
    slide.shapes.add_picture(macro_pic_url, Inches(0.2), Inches(2.1), width=Inches(3.3))
    # 插入右侧：镜下切片图
    slide.shapes.add_picture(slice_pic_url, Inches(4.2), Inches(2.1), width=Inches(3.3))


# ============在这里添加你要复习的病理疾病页面============
add_pathology_slide(
    title="门脉性肝硬化",
    knowledge_text="1.肉眼大体：肝脏缩小、质地变硬，表面布满细小结节\n2.镜下核心特征：假小叶形成，纤维组织广泛增生\n3.考试考点：门脉高压、肝功能损伤两大后果",
    macro_pic_url="这里粘贴网上病理大体图的网址",
    slice_pic_url="这里粘贴肝硬化HE切片的网址"
)

add_pathology_slide(
    title="风湿性心脏病",
    knowledge_text="1.特征病变：风湿小体（Aschoff小体）\n2.瓣膜损伤：二尖瓣出现疣状赘生物\n3.镜下：纤维素样坏死+风湿细胞聚集",
    macro_pic_url="粘贴风湿心脏大体图链接",
    slice_pic_url="粘贴风湿小体切片链接"
)

# 保存最终生成的PPT文件
prs.save("病理复习课件.pptx")
print("PPT生成完成！")