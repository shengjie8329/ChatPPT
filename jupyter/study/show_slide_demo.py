#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
### `Slide` 幻灯片

`Slide` 表示 PowerPoint 演示文稿中的一张幻灯片。

- **属性**：
  - `shapes`：返回该幻灯片中的所有形状，包括文本框、图像、图表等。
  - `slide_layout`：返回该幻灯片的布局对象，指示使用的布局类型。

- **方法**：
  - `shapes.add_shape()`：在t 特定幻灯片上添加一个形状（例如，矩形或圆形）。

- **示例**：

```python
slide = presentation.slides.add_slide(slide_layout)  # 添加一张幻灯片
```

#### `Shape` 形状

- **概述**：`Shape` 表示幻灯片中的一个形状，可以是文本框、图片、图表、SmartArt、表格等。每个 `Shape` 对象都具有位置、大小和样式属性。

- **属性**：
  - `name`: 形状名称，对应 placeholder。
  - `left`、`top`、`width`、`height`：定义形状的位置和尺寸。
  - `text`：如果是文本框，可以访问或修改其内容。

- **方法**：
  - `add_textbox(left, top, width, height)`：用于在幻灯片上添加一个文本框。
  - `add_picture(image_path, left, top, width=None, height=None)`：用于添加图片。

- **示例**：

  ```python
  textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))  # 添加文本框
  ```

@Time   :2024/10/22 15:21
@Author :lancelot.sheng
@File   :show_slide_demo.py
"""
from pptx import Presentation
from pptx.util import Inches

presentation = Presentation("../../outputs/ChatPPT_Demo.pptx")
slide = presentation.slides[0]  # 获取第一张幻灯片
print(slide.shapes[0].text)  # 输出幻灯片中的所有形状

# 打印每页形状 名称和文本，如果是非文本（如 PlaceholderPicture）将会报错
for idx, slide in enumerate(presentation.slides):
    print(f"[slide id]:{idx}")
    for shape in slide.shapes:
        print(f"shape name:{shape.name}")
        print(f"shape text:{shape.text}")
        print("\n")

'''
### 【深入理解】Presentation 和 SlideMaster 类继承关系的 UML 类图

在这个 UML 类图中：

1. **`Presentation`** 类是顶层对象，它包含多个 `Slides` 对象。
2. **`Slides`** 类是一个幻灯片集合，通过它可以添加或访问单独的 `Slide` 对象。
3. **`Slide`** 类代表单个幻灯片，它包含形状（`Shapes`）和占位符（`SlidePlaceholders`），并且它通过布局（`SlideLayout`）来定义外观。
4. **`SlideMaster`** 类包含多个 `SlideLayouts`，它定义了幻灯片的母版布局。
5. **`SlideLayout`** 类定义了幻灯片的布局结构，其中有占位符和形状。
6. **`Shape`** 类代表幻灯片中的形状或文本框等内容。


```
+------------------------+
|      Presentation      |
+------------------------+
| - slides: Slides       |
| - slide_masters: SlideMasters |
| - slide_layouts: SlideLayouts |
| - core_properties: CoreProperties |
+------------------------+
| + save(file: str)                       |
+------------------------+
           |
           v
+--------------------+
|       Slides       |
+--------------------+
| - slides: Slide[]  |
+--------------------+
| + add_slide(layout: SlideLayout) -> Slide |
| + get(slide_id: int) -> Slide | None      |
+--------------------+
           |
           v
+--------------------+
|       Slide        |
+--------------------+
| - slide_id: int    |
| - shapes: Shapes   |
| - placeholders: SlidePlaceholders |
| - slide_layout: SlideLayout       |
+--------------------+
| + add_shape(shape: Shape)         |
| + add_picture(image: Picture)     |
| + add_table(rows: int, cols: int) |
+--------------------+
           |
           v
+--------------------+
|    SlideMaster     |
+--------------------+
| - slide_layouts: SlideLayouts[]  |
+--------------------+
| + get_by_name(name: str) -> SlideLayout |
| + index(slide_layout: SlideLayout) -> int |
+--------------------+
           |
           v
+--------------------+
|   SlideLayouts     |
+--------------------+
| - layouts: SlideLayout[] |
+--------------------+
| + remove(slide_layout: SlideLayout)      |
+--------------------+
           |
           v
+--------------------+
|   SlideLayout      |
+--------------------+
| - placeholders: SlidePlaceholders[] |
| - shapes: Shapes[]                  |
| - slide_master: SlideMaster          |
+--------------------+
           |
           v
+--------------------+
|      Shape         |
+--------------------+
| - name: str        |
| - fill: FillFormat |
| - line: LineFormat |
+--------------------+
| + add_textbox(left, top, width, height)  |
| + add_picture(image_file: str)           |
+--------------------+
```
'''