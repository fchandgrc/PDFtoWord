import pytesseract
from PIL import Image
import docx
import fitz
import os

'''
# 将PDF转化为图片
pdfPath pdf文件的路径
imgPath 图像要保存的文件夹
zoom_x x方向的缩放系数
zoom_y y方向的缩放系数
rotation_angle 旋转角度
'''
def pdf_image(pdfPath,imgPath,zoom_x,zoom_y,rotation_angle):
    # 打开PDF文件
    pdf = fitz.open(pdfPath)
    # 逐页读取PDF
    page_num = pdf.page_count
    for pg in range(0, page_num):
        page = pdf[pg]
        # 设置缩放和旋转系数
        trans = fitz.Matrix(zoom_x, zoom_y)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        # 开始写图像
        pm.save(imgPath+str(pg)+".png")
    pdf.close()
    return page_num
    


if __name__ == '__main__':
    page_num = pdf_image(r"1.pdf",r"C:\\Users\\32878\\Desktop\\code\\pdf\\",5,5,0)

    #创建内存中的word文档对象
    file=docx.Document()
    for ele in range(0, page_num):
        pic_name = str(ele) + ".png"
        image = Image.open(pic_name)
        code = pytesseract.image_to_string(image, lang='eng')
        prev = ""
        strcopy = ""
        for ele in code:
            if ele == 'e' and prev == 'l':
                continue
            if ele == '\n' and prev == '\n':
                strcopy += '\r\n'
            elif ele == '\n' and prev != '\n':
                pass
            else:
                strcopy += ele
            prev = ele
        file.add_paragraph(strcopy)
    file.save("writeResult2.docx")