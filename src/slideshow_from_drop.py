from __future__ import annotations
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


class Rectangle:
    def __init__(self, width, aspect_ratio):
        self.width = width
        self.aspect_ratio = aspect_ratio
        self.height = width / aspect_ratio

    def fit(self, target_rect: Rectangle) -> None:
        """
        Reset width and height to fit the target rectangle.
        Args:
            target_rect: Target Rectangle class instance

        Returns: Rectangle class instance after resize.
        """
        if self.aspect_ratio > target_rect.aspect_ratio:  # 横長
            self.width = target_rect.width
            self.height = self.width / self.aspect_ratio
        else:  # 縦長
            self.height = target_rect.height
            self.width = self.aspect_ratio * self.height

    def fill(self, target_rect: Rectangle) -> None:
        """
        Reset width and height to fill the target rectangle.
        Args:
            target_rect: Target Rectangle class instance

        Returns: Rectangle class instance after resize.

        """
        if self.aspect_ratio > target_rect.aspect_ratio:  # 横長
            self.height = target_rect.height
            self.width = self.aspect_ratio * self.height
        else:  # 縦長
            self.width = target_rect.width
            self.height = self.width / self.aspect_ratio


def add_slide(prs: Presentation()) -> Slides:
    # 白紙スライドの追加(ID=6は白紙スライド)
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    return slide


def add_picture(img_path: str, prs, slide: Slides, grid: tuple, position:int) -> Slides:
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    grid_width, grid_height = (int(slide_width / grid[1]), int(slide_height / grid[0]))
    grid_aspect_ratio = grid_width / grid_height

    # 画像サイズを取得してアスペクト比を得る
    im = Image.open(img_path)
    im_width, im_height = im.size
    im_aspect_ratio = im_width / im_height

    if im_aspect_ratio > grid_aspect_ratio:  # 画像のほうが横長の場合
        img_display_width = grid_width
        img_display_height = img_display_width / im_aspect_ratio
    else:  # 画像のほうが縦長の場合
        img_display_height = grid_height
        img_display_width = img_display_height * im_aspect_ratio
    # グリッドの左上座標を計算
    grid_left = grid_width * (position % grid[1])
    grid_top = grid_height * int(position / grid[1])
    left = grid_left + ((grid_width - img_display_width) / 2)
    top = grid_top + ((grid_height - img_display_height) / 2)

    # 画像をスライドに追加
    slide.shapes.add_picture(img_path, left, top, width=img_display_width)

    return slide


# ファイル名をスライド右上に貼り付ける
def add_filename(file_name: str, prs, slide: Slides, slide_shape: dict):
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    text_box = slide.shapes.add_textbox(0, 0, slide_width, Pt(28))
    # text_box.fill.solid()
    # text_box.fill.fore_color.rgb = RGBColor(0, 0, 0)
    text_frame = text_box.text_frame
    text_frame.text = file_name
    text_frame.paragraphs[0].font.size = Pt(14)
    text_frame.paragraphs[0].font.color.rgb = RGBColor(128, 128, 128)
    text_frame.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.RIGHT

    return slide


def make_slideshow(img_file_paths: list, slide_aspect_ratio: float = 4 / 3, grid: tuple = (1, 1)):
    """
    4:3 (default) 9144000x6858000, 16:9 9144000x5143680
    Unit: English Metric Units = 1/360000 of centimeter
    Width is always 25.4 cm = 25.4 x 360000 EMU = 9144000 EMU.
    Args:
        img_file_paths: list of image file paths (full path)
        slide_aspect_ratio: slide width/height
        grid: image grid numbers as tuple (row, column)

    Returns:
        No returns.

    """

    # region Check Input
    if len(img_file_paths) == 0 or type(img_file_paths) != list:
        print('no image file included. valid extensions: png/jpg/jpeg/bmp only.')
        return
    # endregion

    # region Initialize: Make Slide object, Set output filename
    slide_width = 9144000
    slide_height = int(slide_width / slide_aspect_ratio)  # 6858000
    # Generate presentation object
    prs = Presentation()
    prs.slide_width = slide_width
    prs.slide_height = slide_height
    # Set output params
    output_dir = os.path.dirname(img_file_paths[0])
    print('Output .pptx file dir = ' + output_dir)
    output_file_name = output_dir + '/slideshow_' + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.pptx'
    # Calculate image size
    qty_per_a_slide = int(grid[0])*int(grid[1])
    # endregion

    for i, img_path in enumerate(img_file_paths):
        if i % qty_per_a_slide == 0:
            slide = add_slide(prs)
        position = i % qty_per_a_slide
        add_picture(img_path, prs, slide, grid, position)
        #add_filename(os.path.basename(img_path), slide, {'width': slide_width, 'height': slide_height})

    # save .pptx file
    prs.save(output_file_name)


if __name__ == '__main__':
    img_file_paths = [name for name in sys.argv if name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))]
    #img_file_paths.sort()  # 昇順にsort
    make_slideshow(img_file_paths)
