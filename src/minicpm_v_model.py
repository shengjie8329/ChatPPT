import os
from PIL import Image
from transformers import AutoModel, AutoTokenizer
from logger import LOG  # 引入日志模块，用于记录日志
import base64
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage


def encode_image(image_path):
    """获取图像的 base64 字符串"""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

# os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"
# os.environ["TRANSFORMERS_CACHE"] = "D:\\work\\jk_agent\\hf_cache"

# 加载模型和分词器
# 这里我们使用 `AutoModel` 和 `AutoTokenizer` 加载模型 'openbmb/MiniCPM-V-2_6-int4'
# 参数 `trust_remote_code=True` 表示信任远程代码（根据模型文档设置）

# model = AutoModel.from_pretrained('openbmb/MiniCPM-V-2_6-int4', trust_remote_code=True)
# tokenizer = AutoTokenizer.from_pretrained('openbmb/MiniCPM-V-2_6-int4', trust_remote_code=True)
# model.eval()  # 设置模型为评估模式，以确保不进行训练中的随机性操作

def chat_with_image(image_file, question='简短描述下这幅图，以中文回答，字数不超过100个汉字', sampling=False, temperature=0.7, stream=False):
    """
    使用模型的聊天功能生成对图像的回答。
    
    参数:
        image_file: 图像文件，用于处理的图像。
        question: 提问的问题，默认为 '描述下这幅图'。
        sampling: 是否使用采样进行生成，默认为 False。
        temperature: 采样温度，用于控制生成文本的多样性，值越高生成越多样。
        stream: 是否流式返回响应，默认为 False。
        
    返回:
        生成的回答文本字符串。
    """
    # 打开并转换图像为 RGB 模式
    # image = Image.open(image_file).convert('RGB')

    # 创建消息列表，模拟用户和 AI 的对话
    # msgs = [{'role': 'user', 'content': [image, question]}]

    # 如果不启用流式输出，直接返回生成的完整响应
    # if not stream:
    #     return model.chat(image=None, msgs=msgs, tokenizer=tokenizer, temperature=temperature)
    # else:
    #     # 启用流式输出，则逐字生成并打印响应
    #     generated_text = ""
    #     for new_text in model.chat(image=None, msgs=msgs, tokenizer=tokenizer, sampling=sampling, temperature=temperature, stream=True):
    #         generated_text += new_text
    #         print(new_text, flush=True, end='')  # 实时输出每部分生成的文本
    #     return generated_text  # 返回完整的生成文本
    img_base64 = encode_image(image_file)
    llm = ChatOpenAI(base_url="https://ai-yyds.com/v1", temperature=temperature, model_name="gpt-4o-mini")
    msg = llm.invoke(
        [
            HumanMessage(
                content=[
                    {
                        "type": "text",
                        "text": question,
                    },
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{img_base64}"},
                    },
                ]
            )
        ]
    )
    LOG.info(msg)
    ch_msg = llm.invoke(
        [
            SystemMessage("你是一个翻译专家，擅长将输入的英文翻译成中文，翻译出的内容尽量精练不超过50个汉字，你只需输出对应的翻译内容，不要输出额外信息"),
            HumanMessage(
                f"请翻译以下内容: {msg.content}"
            )
        ]
    )
    LOG.info(ch_msg)
    return ch_msg.content



# 主程序入口
if __name__ == "__main__":
    import sys  # 引入 sys 模块以获取命令行参数
    if len(sys.argv) != 2:
        print("Usage: python src/minicpm_v_model.py <image_file>")  # 提示正确的用法
        sys.exit(1)  # 退出并返回状态码 1，表示错误

    image_file = sys.argv[1]  # 获取命令行传入的图像文件路径
    question = 'What is in the image?'  # 定义默认问题
    response = chat_with_image(image_file, question, sampling=True, temperature=0.7, stream=True)  # 调用生成响应函数
    print("\nFinal Response:", response)  # 输出最终响应
