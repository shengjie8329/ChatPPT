#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
@Time   :2024/11/4 16:10
@Author :lancelot.sheng
@File   :stable_diffusion_api.py
"""
import base64
import datetime
import json
import os

from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate

from logger import LOG  # 导入日志工具
import requests
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage


prompt_file = "./prompts/sd_prompt.txt"
prompt = ''
try:
    with open(prompt_file, "r", encoding="utf-8") as file:
        prompt = file.read().strip()
except FileNotFoundError:
    LOG.error(f"找不到提示文件 {prompt_file}!")
    raise


def generate_sd_prompt(keyword: str) -> str:
    """
    Generate a prompt for stable diffusion based on the given keyword.
    :param keyword:  keyword
    :return:  prompt
    """
    chat_prompt = ChatPromptTemplate.from_messages([
        ("system", prompt),  # 系统提示部分
        ("human", "**用户输入**:\n\n{input}"),  # 消息占位符
    ])
    llm = ChatOpenAI(
        model="gpt-4o-mini",
        temperature=0.7,
        max_tokens=4096,
        base_url="https://ai-yyds.com/v1",
    )
    parser = StrOutputParser()
    chain = chat_prompt | llm | parser
    return chain.invoke({"input": keyword, })


def submit_post(url: str, data: dict):
    """
    Submit a POST request to the given URL with the given data.
    :param url:  url
    :param data: data
    :return:  response
    """
    return requests.post(url, data=json.dumps(data))


def save_encoded_image(b64_image: str, output_path: str) -> str:
    """
    Save the given image to the given output path.
    :param b64_image:  base64 encoded image
    :param output_path:  output path
    :return:  None
    """
    # 判断当前目录下是否存在 output 文件夹，如果不存在则创建
    if not os.path.exists("sd_output"):
        os.mkdir("sd_output")
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    output_path = f"{output_path}_{timestamp}" + ".png"
    # 将文件放入当前目录下的 output 文件夹中
    output_path = f"sd_output/{output_path}"
    with open(output_path, "wb") as f:
        f.write(base64.b64decode(b64_image))
    return output_path


def save_json_file(data: dict, output_path: str):
    """
    Save the given data to the given output path.
    :param data:  data
    :param output_path:  output path
    :return:  None
    """
    # 忽略 data 中的 images 字段
    data.pop('images')
    # 将 data 中的 info 字段转为 json 字符串，info 当前数据需要转义
    data['info'] = json.loads(data['info'])

    # 输出 data.info.infotexts
    info_texts = data['info']['infotexts']
    for info_text in info_texts:
        print(info_text)

    # 判断当前目录下是否存在 output 文件夹，如果不存在则创建
    if not os.path.exists("sd_output"):
        os.mkdir("sd_output")
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    output_path = f"{output_path}_{timestamp}" + ".json"
    # 将文件放入当前目录下的 output 文件夹中
    output_path = f"sd_output/{output_path}"
    with open(output_path, "w") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def generate_image(txt2img_url: str, prompt: str, negative_prompt: str, image_name: str = "", width: int = 1024, height: int = 768):
    """
    Generate an image using the given prompt and negative prompt.
    :param prompt:  prompt
    :param negative_prompt:  negative prompt
    :param image_name:  image name
    :return:  None
    """
    data = {
        'prompt': prompt,
        'negative_prompt': negative_prompt,
        'batch_size': 1,  # 批量大小
        # 'enable_hr': True,  # 是否开启高清
        # 'denoising_strength': 0.5,  # 去噪强度
        'width': width,  # 宽度
        'height': height,  # 高度
        # 'sampler_index': 'Euler a',  # 采样器
        "override_settings": {
            "sd_model_checkpoint": "realisticVisionV20_v20NoVAE.safetensors",
            # "sd_vae": "sdxl_vae.safetensors"
        },
    }
    if not image_name:
        image_name = data['prompt'].replace(" ", "_").replace("/", "_").replace("\\", "_").replace(":", "_").replace(
            "\"","_").replace(
            "<", "_").replace(">", "_").replace("|", "_")[:40]
    sd_response = submit_post(txt2img_url, data)
    path = save_encoded_image(sd_response.json()['images'][0], image_name)
    save_json_file(sd_response.json(), image_name)
    return path


def generate_sd_image(keyword: str, txt2img_url: str, width: int = 1024, height: int = 768 ) ->str:
    """
     Generate a sd image for the given keyword.
    :param keyword:  keyword
    :return:  image_path
    """
    sd_llm_prompt = generate_sd_prompt(keyword)
    LOG.info(f"sd_prompt={sd_llm_prompt}")
    sd_negative_prompt = """(((nsfw))),EasyNegative,badhandv4,ng deepnegative v1 75t(worst quality:2),(low quality:2)(normal quality:2),lowres,((monochrome)),((grayscale)),bad anatomy,DeepNegative,skin spots,acnes,skinblemishes,(fat:1.2),facing away,looking awaytilted head,lowres,bad anatomy,bad hands,missingfingers,extra digit,fewer digits,bad feet,poorly drawn hands,poorly drawn face,mutation,deformed,extrafingers,extra limbs,extra arms,extra legs,malformed limbs,fused fingers,too many fingers,long neck,crosseyed,mutated hands,polar lowres,bad body,bad proportions,gross proportions,missing arms,missing legs,extradigit,extra arms,extra leg,extra foot,teethcroppe,signature,watermark,username,blurry,cropped,jpegartifacts,text,errorLower body exposure,"""
    img_path = generate_image(txt2img_url, sd_llm_prompt, sd_negative_prompt, keyword, width, height)
    LOG.info(f"path={img_path}")
    return img_path




if __name__ == '__main__':
    """
    # txt2img_url = "http://127.0.0.1:7860/sdapi/v1/txt2img" # 服务器地址
    txt2img_url = "http://0or-zbqix1th49192k20c-cwsdao5e-custom.service.onethingrobot.com/sdapi/v1/txt2img"  # 服务器地址
    # prompt = input("请输入提示词：")
   
    # negative_prompt = input("请输入反面提示词：")
    
    data = {
        'prompt': prompt,
        'negative_prompt': negative_prompt,
        'batch_size': 1,  # 批量大小
        # 'enable_hr': True,  # 是否开启高清
        # 'denoising_strength': 0.5,  # 去噪强度
        'width': 1024,  # 第一阶段宽度
        'height': 768,  # 第一阶段高度
        # 'sampler_index': 'Euler a',  # 采样器
        "override_settings": {
            "sd_model_checkpoint": "电商场景MIX_V2.safetensors",
            # "sd_vae": "sdxl_vae.safetensors"
        },
    }
    # 将 data.prompt 中的文本，删除文件名非法字符，已下划线分隔，作为文件名
    output_path = data['prompt'].replace(" ", "_").replace("/", "_").replace("\\", "_").replace(":", "_").replace("\"",
                                                                                                                  "_").replace(
        "<", "_").replace(">", "_").replace("|", "_")[:40]
    response = submit_post(txt2img_url, data)
    save_encoded_image(response.json()['images'][0], output_path)
    save_json_file(response.json(), output_path)
    """

    keyword = "古巴雪茄"
    # sd_prompt = generate_sd_prompt(keyword)
    # print(f"sd_prompt={sd_prompt}")

    txt2img_url = "http://0or-zbqix1th49192k20c-cwsdao5e-custom.service.onethingrobot.com/sdapi/v1/txt2img"
    # prompt = sd_prompt
    # negative_prompt = """(((nsfw))),EasyNegative,badhandv4,ng deepnegative v1 75t(worst quality:2),(low quality:2)(normal quality:2),lowres,((monochrome)),((grayscale)),bad anatomy,DeepNegative,skin spots,acnes,skinblemishes,(fat:1.2),facing away,looking awaytilted head,lowres,bad anatomy,bad hands,missingfingers,extra digit,fewer digits,bad feet,poorly drawn hands,poorly drawn face,mutation,deformed,extrafingers,extra limbs,extra arms,extra legs,malformed limbs,fused fingers,too many fingers,long neck,crosseyed,mutated hands,polar lowres,bad body,bad proportions,gross proportions,missing arms,missing legs,extradigit,extra arms,extra leg,extra foot,teethcroppe,signature,watermark,username,blurry,cropped,jpegartifacts,text,errorLower body exposure,"""
    # path = generate_image(txt2img_url, prompt, negative_prompt, keyword)
    # print(f"path={path}")

    path = generate_sd_image(keyword, txt2img_url)

