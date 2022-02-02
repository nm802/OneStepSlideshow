import os
import sys
import datetime
from PIL import Image
import collections.abc
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor
from pptx.slide import Slides


def add_slide(p: Presentation()) -> Slides:
    # 白紙スライドの追加(ID=6は白紙スライド)
    blank_slide_layout = p.slide_layouts[6]
    slide = p.slides.add_slide(blank_slide_layout)
    return slide


def add_picture(slide: Slides, img_path: str, slide_shape: dict) -> Slides:
    # 画像サイズを取得してアスペクト比を得る
    im = Image.open(img_path)
    im_width, im_height = im.size
    aspect_ratio = im_width / im_height

    slide_aspect_ratio = slide_shape['width']/slide_shape['height']
    if aspect_ratio > slide_aspect_ratio:  # 画像のほうが横長の場合
        img_display_width = slide_shape['width']
        img_display_height = img_display_width / aspect_ratio
    else:  # 画像のほうが縦長の場合
        img_display_height = slide_shape['height']
        img_display_width = img_display_height * aspect_ratio
    # センタリングする場合の画像の左上座標を計算
    img_center_x = slide_shape['width'] / 2
    img_center_y = slide_shape['height'] / 2
    left = img_center_x - img_display_width / 2
    top = img_center_y - img_display_height / 2

    # 画像をスライドに追加
    if aspect_ratio > slide_aspect_ratio:
        slide.shapes.add_picture(img_path, left, top, width=img_display_width)
    else:
        slide.shapes.add_picture(img_path, left, top, height=img_display_height)

    return slide


# ファイル名をスライド右上に貼り付ける
def add_filename(slide: Slides, file_name: str, slide_shape: dict):
    text_box = slide.shapes.add_textbox(0, 0, slide_shape['width'], Pt(28))
    # text_box.fill.solid()
    # text_box.fill.fore_color.rgb = RGBColor(0, 0, 0)
    text_frame = text_box.text_frame
    text_frame.text = file_name
    text_frame.paragraphs[0].font.size = Pt(14)
    text_frame.paragraphs[0].font.color.rgb = RGBColor(128, 128, 128)
    text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.RIGHT

    return slide


def make_slideshow(img_file_paths: list, slide_width: int = 9144000, slide_height: int = 6858000):
    """
    4:3 (default) 9144000x6858000, 16:9 12193200x6858000
    Args:
        slide_width:
        slide_height:

    Returns:

    """

    # スライドオブジェクトの定義
    ppt = Presentation()
    # スライドサイズの指定
    ppt.slide_width = slide_width
    ppt.slide_height = slide_height

    if len(img_file_paths) == 0 or type(img_file_paths) != list:
        print('no image file included. valid extentions: png/jpg/jpeg/bmp only.')
        return

    # 昇順にソート（この順番でスライドに貼り付けられる）
    img_file_paths.sort()  # 昇順にsort
    output_dir = os.path.dirname(img_file_paths[0])
    print('dirname = ' + output_dir)

    for img_path in img_file_paths:
        slide = add_slide(ppt)
        add_picture(slide, img_path, {'width': slide_width, 'height': slide_height})
        add_filename(slide, os.path.basename(img_path), {'width': slide_width, 'height': slide_height})

    # pptxファイルを出力する
    output_file_name = output_dir + '/slideshow_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.pptx'
    print('output file = ' + output_file_name)
    ppt.save(output_file_name)


if __name__ == '__main__':
    img_file_paths = [name for name in sys.argv if name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))]
    make_slideshow(img_file_paths)
