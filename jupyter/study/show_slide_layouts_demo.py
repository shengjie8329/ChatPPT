#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
2. SlideLayout 布局
SlideLayout 表示一个幻灯片的布局，包含一组预定义的占位符和格式设置，定义了幻灯片的基本结构。

获取布局：

presentation.slide_layouts：这是一个布局列表，包含了所有可用的幻灯片布局。

单个布局:

slide_layout = presentation.slide_layouts[0]  # 标题幻灯片布局
slide_layout = presentation.slide_layouts[1]  # 内容幻灯片

@Time   :2024/10/22 14:00
@Author :lancelot.sheng
@File   :show_slide_layouts_demo.py
"""
from pptx import Presentation
from pptx.util import Inches

template = Presentation("../../templates/MasterTemplate.pptx")
slide_layouts = template.slide_layouts
print(slide_layouts)

if len(slide_layouts) > 0:
    slide_layout = template.slide_layouts[0]
    print(slide_layout)
    print(slide_layout.name)
    # 打印所有的shape
    for shape in slide_layout.shapes:
        print(" "+shape.name)

    # 打印所有的slide_layout
    print("---------所有的slide_layout")
    for s in slide_layouts:
        print(s.name)