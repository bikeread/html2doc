"""
测试HTML到DOCX的完整转换过程
"""
import logging
import os
from converters.docx_converter import DocxConverter
import datetime

# 配置日志
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# 测试HTML内容
TEST_HTML = """
<!DOCTYPE html>
<html>
<head>
    <style>
        /* 论文封面 */
        .cover {
            font-family: "SimHei";
            font-size: 17pt;
        }

        /* 目录 */
        .toc {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }

        /* 摘要 */
        .abstract h1,
        .abstract .summary-title {
            font-family: "SimHei";
            font-size: 16pt;
            text-align: center;
        }
        
        .abstract p,
        .abstract .summary-content {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }

        /* 关键词 */
        .keywords {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }
        
        .keywords .keyword-title {
            font-weight: bold;
        }

        /* 正文标题 */
        .chapter-title {
            font-family: "SimHei";
            font-size: 16pt;
            text-align: center;
            margin-top: 1em;
            line-height: 1.5;
        }
        
        .section-title {
            font-family: "SimHei";
            font-size: 14pt;
            margin-top: 0.5em;
            line-height: 1.5;
        }
        
        .subsection-title {
            font-family: "SimSun";
            font-size: 12pt;
            font-weight: bold;
            line-height: 1.5;
        }

        /* 正文 */
        body {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }
        
        p.english {
            font-family: "Times New Roman";
            font-size: 12pt;
            line-height: 1.5;
        }

        /* 注释 */
        .notes {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }

        /* 参考文献 */
        .references {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }

        /* 谢辞(或致谢) */
        .acknowledgments {
            font-family: "SimSun";
            font-size: 12pt;
            line-height: 1.5;
        }
    </style>
</head>
<body>
    <!-- 论文封面 -->
    <div class="cover">
        <h1>论文标题</h1>
        <p>作者姓名</p>
        <p>学校/机构名称</p>
        <p>日期：2025年3月</p>
    </div>

    <!-- 目录 -->
    <div class="toc">
        <h2>目录</h2>
        <p>1. 第一章 <span style="float:right">1</span></p>
        <p>&nbsp;&nbsp;1.1 第一节 <span style="float:right">2</span></p>
        <p>&nbsp;&nbsp;1.2 第二节 <span style="float:right">3</span></p>
        <p>2. 第二章 <span style="float:right">4</span></p>
        <p>&nbsp;&nbsp;2.1 第一节 <span style="float:right">5</span></p>
    </div>

    <!-- 摘要 -->
    <div class="abstract">
        <h1>摘要</h1>
        <p class="summary-content">本研究旨在探讨某领域的重要问题，通过系统分析和实验验证，提出了创新性解决方案。研究结果表明，所提出的方法在多个方面优于现有技术，为该领域的发展提供了新的思路和方向。</p>
    </div>

    <!-- 关键词 -->
    <div class="keywords">
        <p><strong class="keyword-title">关键词：</strong>人工智能, 机器学习, 数据挖掘, 深度学习, 知识图谱</p>
    </div>

    <!-- 正文章节 -->
    <h1 class="chapter-title">第一章 绪论</h1>
    
    <h2 class="section-title">1.1 研究背景</h2>
    <p>随着信息技术的快速发展，人工智能已经成为推动社会进步的重要力量。本研究从实际应用需求出发，探索人工智能在特定领域的应用价值和实现路径。</p>
    
    <h3 class="subsection-title">1.1.1 国内外研究现状</h3>
    <p>近年来，国内外学者在该领域进行了大量研究，取得了丰硕成果。国外研究主要集中在理论创新和算法优化方面，而国内研究则更注重应用落地和产业化。</p>
    
    <p class="english">This paper presents a novel approach to artificial intelligence applications in specific domains, with emphasis on practical implementations and theoretical foundations.</p>
    
    <h2 class="section-title">1.2 研究目的与意义</h2>
    <p>本研究旨在解决实际应用中面临的关键问题，提高系统效率和准确性，为相关领域的发展提供理论支撑和技术方案。</p>

    <h1 class="chapter-title">第二章 理论基础</h1>
    
    <h2 class="section-title">2.1 基本概念</h2>
    <p>本章介绍研究所涉及的基本理论和核心概念，为后续研究奠定基础。</p>

    <!-- 注释 -->
    <div class="notes">
        <h2>注释</h2>
        <p>[1] 这里对特定术语或概念进行补充说明。</p>
        <p>[2] 提供额外的背景信息或解释。</p>
    </div>

    <!-- 参考文献 -->
    <div class="references">
        <h2>参考文献</h2>
        <p>[1] 作者. 论文题目[J]. 期刊名称, 年份, 卷(期): 起止页码.</p>
        <p>[2] 作者. 书名[M]. 出版地: 出版社, 出版年份.</p>
        <p>[3] 作者. 网站名称[EB/OL]. 网址, 访问日期.</p>
    </div>

    <!-- 致谢 -->
    <div class="acknowledgments">
        <h2>致谢</h2>
        <p>感谢导师的悉心指导和同学们的大力支持，感谢家人的理解和鼓励。本研究得到了某某基金项目（编号：XXXX）的资助。</p>
    </div>
</body>
</html>
"""


def main():
    """测试HTML到DOCX的完整转换"""
    html_file = "test_input.html"
    
    # 生成带时间戳的输出文件名
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"test_output_{timestamp}.docx"
    
    print(f"正在转换 {html_file} 到 {output_file}...")
    
    # 从文件读取HTML内容
    try:
        with open(html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except FileNotFoundError:
        # 如果找不到文件，使用测试用的HTML内容
        print(f"未找到文件 {html_file}，使用内置的测试HTML")
        html_content = TEST_HTML
    
    # 创建转换器实例并执行转换
    converter = DocxConverter()
    result = converter.convert(html_content)
    
    # 保存转换结果
    with open(output_file, "wb") as f:
        f.write(result)
    
    print(f"转换结果已保存到: {os.path.abspath(output_file)}")


if __name__ == "__main__":
    main()
