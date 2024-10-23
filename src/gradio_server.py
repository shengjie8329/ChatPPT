#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@Time   :2024/10/23 16:30
@Author :lancelot.sheng
@File   :gradio_server.py
"""
import gradio as gr
from conversation_agent import ConversationAgent
from input_parser import parse_input_text
from ppt_generator import generate_presentation
from template_manager import load_template, get_layout_mapping, print_layouts
from layout_manager import LayoutManager
from logger import LOG

conversation_agent = ConversationAgent(session_id="test_user")

# message = conversation_agent.chat_with_history("请写一份简单的周报，内容随意，事项不要超过5条")
#
# print(message)

templates_info = {
    "MasterTemplate": {
        "layout_mapping": {
            "Title Only": 0,
            "Title and Content": 1,
            "Title and Picture": 2,
            "Title, Content, and Picture": 3
        }
    },
    "LGBTTemplate": {
        "layout_mapping": {
            "Title Only": 0,
            "Title and Content": 1,
            "Title and Picture": 2,
            "Title, Content, and Picture": 3
        }
    }

}


def create_scenario_tab():
    with gr.Tab("场景"):  # 场景标签
        gr.Markdown("## 选择一个场景完成目标和挑战")  # 场景选择说明

        # 创建单选框组件
        scenario_radio = gr.Radio(
            choices=[
                ("模板选择", "MasterTemplate"),
                ("LGBT", "LGBTTemplate"),
            ],
            label="模板选择"  # 单选框标签
        )


def generate_ppt(templt, taskInfo):
    LOG.debug(f"[翻译任务]\n模板: {templt}\n任务描述: {taskInfo}\n")

    mapping_dict = templates_info[templt]["layout_mapping"]
    LOG.debug(f"[模板映射]\n: {mapping_dict}")

    task_markdown = conversation_agent.chat_with_history(taskInfo)

    LOG.debug(f"[ppt markdown]\n: {task_markdown}")

    template_path = f"templates/{templt}.pptx"
    LOG.debug(f"[ppt template_path]\n: {template_path}")

    # 加载 PowerPoint 模板，并打印模板中的可用布局
    prs = load_template(template_path)  # 加载模板文件
    LOG.info("可用的幻灯片布局:")  # 记录信息日志，打印可用布局
    print_layouts(prs)  # 打印模板中的布局

    # 初始化 LayoutManager，使用配置文件中的 layout_mapping
    layout_manager = LayoutManager(mapping_dict)

    # 调用 parse_input_text 函数，解析输入文本，生成 PowerPoint 数据结构
    powerpoint_data, presentation_title = parse_input_text(task_markdown, layout_manager)

    LOG.info(f"解析转换后的 ChatPPT PowerPoint 数据结构:\n{powerpoint_data}")  # 记录调试日志，打印解析后的 PowerPoint 数据
    # 定义输出 PowerPoint 文件的路径
    output_pptx = f"outputs/{presentation_title}.pptx"

    # 调用 generate_presentation 函数生成 PowerPoint 演示文稿
    generate_presentation(powerpoint_data, template_path, output_pptx)
    return output_pptx


def launch_gradio():
    iface = gr.Interface(
        fn=generate_ppt,
        title="PPT生成器 v0.2",
        inputs=[
            gr.Radio(
                choices=[
                    ("Master", "MasterTemplate"),
                    ("LGBT", "LGBTTemplate"),
                ],
                label="模板选择"  # 单选框标签
            ),
            gr.Textbox(label="生成任务描述", placeholder="请写一份简单的周报，内容随意，事项不要超过5条"),
        ],
        outputs=[
            gr.File(label="下载PPT文件")
        ],
        allow_flagging="never"
    )

    iface.launch(share=True, server_name="0.0.0.0")


if __name__ == "__main__":
    # message = conversation_agent.chat_with_history("请写一份简单的周报，内容随意，至少要5条slide")
    #
    # print(message)

    # 启动 Gradio 服务
    launch_gradio()