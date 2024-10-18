import os
from input_parser import parse_input_text
from ppt_generator import generate_presentation
from template_manager import load_template, get_layout_mapping, print_layouts


def main():
    input_text = """
    # ChatPPT_Demo

    ## ChatPPT Demo [Title Only]

    ## 2024 业绩概述 [Title and Content]
    - 总收入增长15%
    - 市场份额扩大至30%

    ## 业绩图表 [Title and Picture 1]
    ![业绩图表](images/performance_chart.png)

    ## 新产品发布 [Title and 2 Column]
    - 产品A: 特色功能介绍
    - 产品B: 市场定位
    ![未来增长](images/forecast.png)
    """

    input_text_lgbt = """
        # ChatPPT_LGBT

        ## ChatPPT_LGBT Demo [Title]

        ## 2024 业绩概述 [Right Pattern Content]
        - 总收入增长15%
        - 市场份额扩大至30%

        ## 业绩图表 [Overview]
        ![业绩图表](images/performance_chart.png)

        ## 新产品发布 [Chart Slide]
        - 产品A: 特色功能介绍
        - 产品B: 市场定位
        ![未来增长](images/forecast.png)
        
        ## 2024 业绩概述1 [Left Pattern Content]
        - 总收入增长15%
        - 市场份额扩大至30%
        
        ## 2024 业绩概述2 [Smart Art]
        - 总收入增长15%
        - 市场份额扩大至30%
        
        ## 2024 业绩概述3 [Two Photo Content]
        - 总收入增长15%
        - 市场份额扩大至30%
        
        ## 2024 业绩概述4 [Right Pattern Content Blue title]
        - 总收入增长15%
        - 市场份额扩大至30%
        
        ## 2024 业绩概述5 [Questions]
        - 总收入增长15%
        - 市场份额扩大至30%
        
        ## 2024 业绩概述6 [Left Pattern Content Orange Title]
        - 总收入增长15%
        - 市场份额扩大至30%
        """

    template_file = 'templates/LGBTTemplate.pptx'
    prs = load_template(template_file)

    print("Available Slide Layouts:")
    print_layouts(prs)

    layout_mapping = get_layout_mapping(prs)

    powerpoint_data, presentation_title = parse_input_text(input_text_lgbt, layout_mapping)

    output_pptx = f"outputs/{presentation_title}.pptx"
    generate_presentation(powerpoint_data, template_file, output_pptx)


if __name__ == "__main__":
    main()
