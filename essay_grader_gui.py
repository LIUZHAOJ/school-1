# -*- coding: utf-8 -*-
"""
小学语文作文智能批改助手 v5.3
基于 DeepSeek API + RapidOCR，支持文字输入、图片识别、Word拖拽、成绩查询
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from openai import OpenAI
import json, time, os, threading, sys, glob

# 尝试导入tkinterdnd2以支持拖放功能
try:
    from tkinterdnd2 import TkinterDnD
    TKINTER_DND_AVAILABLE = True
except ImportError:
    TKINTER_DND_AVAILABLE = False

# ====================== 配置区 ======================
DEEPSEEK_API_KEY = "sk-b2daed9e6e4b47e89d066608069b2e19"
MODEL_NAME = "deepseek-chat"

# ====================== OCR ======================
_ocr = None
def get_ocr():
    global _ocr
    if _ocr is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _ocr = RapidOCR()
        except ImportError:
            raise ImportError("请运行: pip install rapidocr_onnxruntime")
    return _ocr

def ocr_image(path):
    ocr = get_ocr()
    result, _ = ocr(path)
    if not result:
        return ""
    return "\n".join([item[1] for item in result])

# ====================== Word文档读取 ======================
def read_word_text(file_path):
    """读取Word文档内容"""
    try:
        from docx import Document
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        return text
    except ImportError:
        raise ImportError("请运行: pip install python-docx")
    except Exception as e:
        raise Exception(f"读取Word文档失败: {str(e)}")

# ====================== 批改API ======================
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url='https://api.deepseek.com')

SYSTEM_PROMPT = """你是一位拥有20年教学经验的资深小学语文教师，正在批改三年级学生作文。
先鼓励再指出问题。检查错别字、标点、句子通顺、段落结构、内容主题。
输出严格JSON格式：{
    "student_name":"", "essay_title":"", "total_score":0,
    "dimensions":{
        "wrong_characters":{"score":0,"issues":[],"comment":""},
        "punctuation":{"score":0,"issues":[],"comment":""},
        "fluency":{"score":0,"issues":[],"comment":""},
        "structure":{"score":0,"has_intro":true,"has_body":true,"has_ending":true,"comment":""},
        "content":{"score":0,"on_topic":true,"comment":""}
    },
    "overall_comment":"", "highlight_good":[], "improvement_tip":""
}"""

def grade_essay_api(essay_text, student_name, essay_title):
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"学生姓名：{student_name}\n作文题目：{essay_title}\n作文正文：\n{essay_text}\n请严格按JSON格式输出批改结果。"}
        ],
        temperature=0.3, max_tokens=2000
    )
    text = response.choices[0].message.content.strip()
    if text.startswith("```json"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return json.loads(text.strip(), strict=False), response.usage.total_tokens

# ====================== 学生报告生成 ======================
def generate_student_report(result, output_dir=None):
    """生成学生专属HTML报告"""
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "学生报告")
    os.makedirs(output_dir, exist_ok=True)
    
    student_name = result.get('student_name', '学生')
    essay_title = result.get('essay_title', '未知题目')
    total_score = result.get('total_score', 0)
    dims = result.get('dimensions', {})
    highlights = result.get('highlight_good', [])
    improvement = result.get('improvement_tip', '')
    comment = result.get('overall_comment', '暂无评语')
    
    # 评分颜色
    if total_score >= 90:
        score_color = "#27AE60"
        score_emoji = "🌟"
    elif total_score >= 80:
        score_color = "#F39C12"
        score_emoji = "😊"
    elif total_score >= 70:
        score_color = "#E67E22"
        score_emoji = "👍"
    else:
        score_color = "#E74C3C"
        score_emoji = "💪"
    
    # 各维度得分
    wrong_chars = dims.get('wrong_characters', {}).get('score', 0)
    punctuation = dims.get('punctuation', {}).get('score', 0)
    fluency = dims.get('fluency', {}).get('score', 0)
    structure = dims.get('structure', {}).get('score', 0)
    content = dims.get('content', {}).get('score', 0)
    
    # 亮点列表HTML
    highlights_html = ""
    if highlights:
        for h in highlights:
            highlights_html += f'<li>✨ {h}</li>'
    else:
        highlights_html = '<li>暂无特别亮点，继续加油！</li>'
    
    # 改进建议HTML
    improvement_html = f'<p>{improvement}</p>' if improvement else '<p>暂无特别建议，继续保持！</p>'
    
    # 生成文件名（去掉特殊字符）
    safe_name = student_name.replace('/', '_').replace('\\', '_').replace(' ', '_')
    safe_title = essay_title.replace('/', '_').replace('\\', '_').replace(' ', '_')
    timestamp = time.strftime('%Y%m%d')
    filename = f"{safe_name}_{safe_title}_{timestamp}.html"
    filepath = os.path.join(output_dir, filename)
    
    # HTML内容
    html_content = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{student_name}的作文报告 - {essay_title}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ 
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        .container {{ max-width: 700px; margin: 0 auto; }}
        .card {{
            background: white;
            border-radius: 20px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
            margin-bottom: 20px;
        }}
        .header {{
            background: linear-gradient(135deg, #185FA5, #2E86DE);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{ font-size: 24px; margin-bottom: 10px; }}
        .header .info {{ opacity: 0.9; font-size: 14px; }}
        .score-section {{
            text-align: center;
            padding: 40px;
            background: white;
        }}
        .big-score {{
            font-size: 80px;
            font-weight: bold;
            color: {score_color};
            margin: 10px 0;
        }}
        .score-emoji {{ font-size: 50px; }}
        .score-label {{ 
            color: #666;
            font-size: 16px;
            text-transform: uppercase;
            letter-spacing: 2px;
        }}
        .dimensions {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            padding: 20px;
            background: #f8f9fa;
        }}
        .dim-item {{
            background: white;
            border-radius: 10px;
            padding: 15px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }}
        .dim-label {{ color: #666; font-size: 12px; margin-bottom: 5px; }}
        .dim-score {{ font-size: 24px; font-weight: bold; color: #185FA5; }}
        .section {{
            padding: 25px 30px;
            border-top: 1px solid #eee;
        }}
        .section h2 {{
            color: #185FA5;
            font-size: 16px;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .highlights ul {{
            list-style: none;
            padding: 0;
        }}
        .highlights li {{
            padding: 10px 15px;
            margin-bottom: 8px;
            background: #E8F6EF;
            border-left: 4px solid #27AE60;
            border-radius: 0 8px 8px 0;
            color: #2C5F2D;
        }}
        .improvement p {{
            padding: 15px;
            background: #FEF9E7;
            border-left: 4px solid #F39C12;
            border-radius: 0 8px 8px 0;
            color: #9A7B0A;
        }}
        .comment-box {{
            background: #F0F8FF;
            padding: 20px;
            border-radius: 10px;
            line-height: 1.8;
            color: #333;
        }}
        .footer {{
            text-align: center;
            padding: 20px;
            color: rgba(255,255,255,0.7);
            font-size: 12px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <div class="header">
                <h1>{score_emoji} {student_name}的作文报告 {score_emoji}</h1>
                <div class="info">
                    <p>📝 作文题目：{essay_title}</p>
                    <p>📅 批改时间：{time.strftime('%Y年%m月%d日 %H:%M')}</p>
                </div>
            </div>
            
            <div class="score-section">
                <div class="score-label">作 文 总 分</div>
                <div class="big-score">{total_score}</div>
                <div class="score-emoji">{score_emoji}</div>
            </div>
            
            <div class="dimensions">
                <div class="dim-item">
                    <div class="dim-label">📚 字词</div>
                    <div class="dim-score">{wrong_chars}</div>
                </div>
                <div class="dim-item">
                    <div class="dim-label">✨ 标点</div>
                    <div class="dim-score">{punctuation}</div>
                </div>
                <div class="dim-item">
                    <div class="dim-label">🔄 流畅</div>
                    <div class="dim-score">{fluency}</div>
                </div>
                <div class="dim-item">
                    <div class="dim-label">📐 结构</div>
                    <div class="dim-score">{structure}</div>
                </div>
                <div class="dim-item">
                    <div class="dim-label">📖 内容</div>
                    <div class="dim-score">{content}</div>
                </div>
            </div>
            
            <div class="section highlights">
                <h2>🌟 写得好的地方</h2>
                <ul>{highlights_html}</ul>
            </div>
            
            <div class="section improvement">
                <h2>💡 需要改进的地方</h2>
                {improvement_html}
            </div>
            
            <div class="section">
                <h2>📝 老师的话</h2>
                <div class="comment-box">{comment}</div>
            </div>
        </div>
        
        <div class="footer">
            <p>由 小学语文作文智能批改助手 v5.3 生成</p>
        </div>
    </div>
</body>
</html>'''
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    return filepath


# ====================== 班级报告生成 ======================
def generate_class_report(output_dir=None):
    """生成班级作文分析报告（HTML）"""
    if output_dir is None:
        result_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
    else:
        result_dir = output_dir
    
    # 读取所有批改结果
    records = []
    if os.path.exists(result_dir):
        for f in glob.glob(os.path.join(result_dir, "*.json")):
            try:
                with open(f, 'r', encoding='utf-8') as fp:
                    data = json.load(fp)
                    records.append(data)
            except:
                pass
    
    if not records:
        return None, "暂无批改记录"
    
    # 统计分析
    total_students = len(records)
    scores = [r.get('total_score', 0) for r in records]
    avg_score = sum(scores) / len(scores) if scores else 0
    max_score = max(scores) if scores else 0
    min_score = min(scores) if scores else 0
    
    # 各维度平均分
    dims_list = [r.get('dimensions', {}) for r in records]
    avg_dims = {
        '字词': sum([d.get('wrong_characters', {}).get('score', 0) for d in dims_list]) / total_students,
        '标点': sum([d.get('punctuation', {}).get('score', 0) for d in dims_list]) / total_students,
        '流畅': sum([d.get('fluency', {}).get('score', 0) for d in dims_list]) / total_students,
        '结构': sum([d.get('structure', {}).get('score', 0) for d in dims_list]) / total_students,
        '内容': sum([d.get('content', {}).get('score', 0) for d in dims_list]) / total_students,
    }
    
    # 分数段分布
    score_ranges = {'90-100': 0, '80-89': 0, '70-79': 0, '60-69': 0, '60以下': 0}
    for s in scores:
        if s >= 90:
            score_ranges['90-100'] += 1
        elif s >= 80:
            score_ranges['80-89'] += 1
        elif s >= 70:
            score_ranges['70-79'] += 1
        elif s >= 60:
            score_ranges['60-69'] += 1
        else:
            score_ranges['60以下'] += 1
    
    # 找出最高分和最低分学生
    sorted_records = sorted(records, key=lambda x: x.get('total_score', 0), reverse=True)
    top_student = sorted_records[0] if sorted_records else None
    bottom_student = sorted_records[-1] if sorted_records else None
    
    # 找出最常见的问题
    all_improvements = []
    for r in records:
        tip = r.get('improvement_tip', '')
        if tip:
            all_improvements.append(tip)
    
    # 生成文件名
    timestamp = time.strftime('%Y%m%d_%H%M')
    filename = f"班级作文分析报告_{timestamp}.html"
    report_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "班级报告")
    os.makedirs(report_dir, exist_ok=True)
    filepath = os.path.join(report_dir, filename)
    
    # 生成维度柱状图的HTML
    dim_bars = ""
    for dim_name, dim_score in avg_dims.items():
        bar_width = (dim_score / 30) * 100  # 假设满分30
        dim_bars += f'''
        <div style="margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                <span>{dim_name}</span>
                <span>{dim_score:.1f}分</span>
            </div>
            <div style="background: #e0e0e0; border-radius: 10px; height: 20px; overflow: hidden;">
                <div style="background: linear-gradient(90deg, #185FA5, #2E86DE); height: 100%; width: {bar_width}%; border-radius: 10px;"></div>
            </div>
        </div>'''
    
    # 分数段柱状图
    range_bars = ""
    max_count = max(score_ranges.values()) if score_ranges.values() else 1
    for range_name, count in score_ranges.items():
        bar_width = (count / max_count) * 100 if max_count > 0 else 0
        range_bars += f'''
        <div style="margin-bottom: 12px;">
            <div style="display: flex; justify-content: space-between; margin-bottom: 3px;">
                <span>{range_name}分</span>
                <span>{count}人</span>
            </div>
            <div style="background: #e0e0e0; border-radius: 6px; height: 16px; overflow: hidden;">
                <div style="background: linear-gradient(90deg, #27AE60, #2ECC71); height: 100%; width: {bar_width}%; border-radius: 6px;"></div>
            </div>
        </div>'''
    
    # 学生列表
    student_list = ""
    for i, r in enumerate(sorted_records, 1):
        s_score = r.get('total_score', 0)
        if s_score >= 90:
            s_color = "#27AE60"
        elif s_score >= 80:
            s_color = "#F39C12"
        elif s_score >= 70:
            s_color = "#E67E22"
        else:
            s_color = "#E74C3C"
        student_list += f'''
        <tr>
            <td>{i}</td>
            <td>{r.get('student_name', '未知')}</td>
            <td>{r.get('essay_title', '未知')}</td>
            <td style="color: {s_color}; font-weight: bold;">{s_score}</td>
        </tr>'''
    
    # HTML内容
    html_content = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>班级作文分析报告</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ 
            font-family: "Microsoft YaHei", "PingFang SC", sans-serif; 
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            min-height: 100vh;
            padding: 20px;
            color: #333;
        }}
        .container {{ max-width: 1000px; margin: 0 auto; }}
        .header {{
            background: linear-gradient(135deg, #185FA5, #2E86DE);
            color: white;
            padding: 30px;
            border-radius: 20px 20px 0 0;
            text-align: center;
        }}
        .header h1 {{ font-size: 28px; margin-bottom: 10px; }}
        .header .subtitle {{ opacity: 0.9; }}
        .card {{
            background: white;
            border-radius: 0 0 20px 20px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            padding: 30px;
            margin-bottom: 20px;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-box {{
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            padding: 20px;
            border-radius: 15px;
            text-align: center;
        }}
        .stat-box.green {{ background: linear-gradient(135deg, #11998e, #38ef7d); }}
        .stat-box.orange {{ background: linear-gradient(135deg, #f093fb, #f5576c); }}
        .stat-box.blue {{ background: linear-gradient(135deg, #4facfe, #00f2fe); }}
        .stat-value {{ font-size: 36px; font-weight: bold; }}
        .stat-label {{ opacity: 0.9; font-size: 14px; }}
        .two-col {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
        }}
        .section-title {{
            color: #185FA5;
            font-size: 18px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #185FA5;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }}
        th, td {{
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #eee;
        }}
        th {{ background: #f0f4f8; color: #185FA5; }}
        .footer {{
            text-align: center;
            color: rgba(255,255,255,0.6);
            padding: 20px;
            font-size: 12px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 班级作文分析报告</h1>
            <div class="subtitle">生成时间：{time.strftime('%Y年%m月%d日 %H:%M')} | 共 {total_students} 名学生</div>
        </div>
        
        <div class="card">
            <div class="stats-grid">
                <div class="stat-box">
                    <div class="stat-value">{total_students}</div>
                    <div class="stat-label">批改人数</div>
                </div>
                <div class="stat-box green">
                    <div class="stat-value">{avg_score:.1f}</div>
                    <div class="stat-label">班级平均分</div>
                </div>
                <div class="stat-box orange">
                    <div class="stat-value">{max_score}</div>
                    <div class="stat-label">最高分</div>
                </div>
                <div class="stat-box blue">
                    <div class="stat-value">{min_score}</div>
                    <div class="stat-label">最低分</div>
                </div>
            </div>
            
            <div class="two-col">
                <div>
                    <h3 class="section-title">📈 各维度平均得分</h3>
                    {dim_bars}
                </div>
                <div>
                    <h3 class="section-title">📊 分数段分布</h3>
                    {range_bars}
                </div>
            </div>
        </div>
        
        <div class="card">
            <h3 class="section-title">🏆 学生成绩排名</h3>
            <table>
                <thead>
                    <tr>
                        <th>排名</th>
                        <th>姓名</th>
                        <th>作文题目</th>
                        <th>总分</th>
                    </tr>
                </thead>
                <tbody>
                    {student_list}
                </tbody>
            </table>
        </div>
        
        <div class="card">
            <h3 class="section-title">⭐ 优秀学生</h3>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                <div style="background: #E8F6EF; padding: 15px; border-radius: 10px;">
                    <div style="color: #27AE60; font-weight: bold; margin-bottom: 5px;">🌟 最高分</div>
                    <div>{top_student.get('student_name', '未知')} - {top_student.get('total_score', 0)}分</div>
                    <div style="font-size: 12px; color: #666;">{top_student.get('essay_title', '')}</div>
                </div>
                <div style="background: #FEF9E7; padding: 15px; border-radius: 10px;">
                    <div style="color: #F39C12; font-weight: bold; margin-bottom: 5px;">💪 进步空间</div>
                    <div>{bottom_student.get('student_name', '未知')} - {bottom_student.get('total_score', 0)}分</div>
                    <div style="font-size: 12px; color: #666;">{bottom_student.get('improvement_tip', '')[:50]}...</div>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p>由 小学语文作文智能批改助手 v5.3 生成 | 仅供教师内部参考</p>
        </div>
    </div>
</body>
</html>'''
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    return filepath, f"共分析 {total_students} 名学生的作文"


# ====================== GUI主类 ======================
class EssayGraderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小学语文作文智能批改助手 v5.3")
        # 增大窗口尺寸，让所有内容一目了然
        self.root.geometry("1500x900")
        self.root.configure(bg="#f0f4f8")
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"1500x900+{(sw-1500)//2}+{(sh-900)//2}")

        self.ocr_path = None
        self.batch_files = []
        self.batch_results = []
        self.history_data = []
        self.filtered_history = []
        self.setup_ui()

    def setup_ui(self):
        # --- 顶部标题栏（精简高度）---
        title_frame = tk.Frame(self.root, bg="#185FA5", height=50)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        
        # 标题文字
        title_left = tk.Frame(title_frame, bg="#185FA5")
        title_left.pack(side="left", padx=20, pady=6)
        tk.Label(title_left, text="📝", font=("Microsoft YaHei", 18), fg="white", bg="#185FA5").pack(side="left", padx=(0, 8))
        title_text_frame = tk.Frame(title_left, bg="#185FA5")
        title_text_frame.pack(side="left")
        tk.Label(title_text_frame, text="小学语文作文智能批改助手 v5.3", font=("Microsoft YaHei", 13, "bold"), fg="white", bg="#185FA5").pack(anchor="w")
        tk.Label(title_text_frame, text="DeepSeek API + RapidOCR离线识别 · 智能五维评分", font=("Microsoft YaHei", 8), fg="#B5D4F4", bg="#185FA5").pack(anchor="w")
        
        # 右侧状态
        title_right = tk.Frame(title_frame, bg="#185FA5")
        title_right.pack(side="right", padx=20, pady=6)
        self.title_status = tk.Label(title_right, text="✅ 就绪", font=("Microsoft YaHei", 9), fg="#90EE90", bg="#185FA5")
        self.title_status.pack()
        
        # --- 主内容区（左右分栏）---
        main_paned = ttk.PanedWindow(self.root, orient="horizontal")
        main_paned.pack(fill="both", expand=True, padx=8, pady=8)
        
        # ===== 左侧：输入区 =====
        left_frame = tk.Frame(main_paned, bg="#f5f7fa")
        main_paned.add(left_frame, weight=3)  # 左侧占60%
        
        # --- 顶部信息区（学生姓名+题目）---
        info_card = tk.Frame(left_frame, bg="white", padx=12, pady=10)
        info_card.pack(fill="x", padx=5, pady=(0, 5))
        
        # 第一行：学生姓名和作文题目
        info_row1 = tk.Frame(info_card, bg="white")
        info_row1.pack(fill="x", pady=(0, 6))
        
        tk.Label(info_row1, text="学生姓名：", font=("Microsoft YaHei", 10), bg="white", width=10, anchor="e").pack(side="left", padx=(0, 5))
        self.name_entry = tk.Entry(info_row1, font=("Microsoft YaHei", 11), bd=1, relief="solid")
        self.name_entry.pack(side="left", fill="x", expand=True, padx=(0, 15))
        
        tk.Label(info_row1, text="作文题目：", font=("Microsoft YaHei", 10), bg="white", width=10, anchor="e").pack(side="left", padx=(0, 5))
        self.title_entry = tk.Entry(info_row1, font=("Microsoft YaHei", 11), bd=1, relief="solid")
        self.title_entry.pack(side="left", fill="x", expand=True)
        
        # --- 醒目的主按钮区（开始批改）---
        grade_card = tk.Frame(left_frame, bg="white", padx=12, pady=12)
        grade_card.pack(fill="x", padx=5, pady=(0, 5))
        
        # 主按钮：开始批改（超大醒目按钮）
        self.grade_btn = tk.Button(grade_card, text="🚀 开始批改", font=("Microsoft YaHei", 15, "bold"), fg="white", bg="#27AE60", activebackground="#1E8449", relief="flat", padx=30, pady=12, cursor="hand2", command=self.on_grade)
        self.grade_btn.pack(fill="x", padx=(0, 0))
        
        # 次要按钮行
        btn_row = tk.Frame(grade_card, bg="white")
        btn_row.pack(fill="x", pady=(8, 0))
        
        self.demo_btn = tk.Button(btn_row, text="📄 示例", font=("Microsoft YaHei", 10), fg="#3B6D11", bg="white", relief="solid", bd=1, padx=12, pady=5, cursor="hand2", command=self.load_demo)
        self.demo_btn.pack(side="left", padx=(0, 8))
        self.clear_btn = tk.Button(btn_row, text="🗑 清空", font=("Microsoft YaHei", 10), fg="#888888", bg="white", relief="solid", bd=1, padx=12, pady=5, cursor="hand2", command=self.on_clear)
        self.clear_btn.pack(side="left")
        
        self.status_label = tk.Label(btn_row, text="", font=("Microsoft YaHei", 9), fg="#888888", bg="white")
        self.status_label.pack(side="right")
        
        # --- 作文正文输入区（Tab页）---
        input_card = tk.Frame(left_frame, bg="white", padx=12, pady=10)
        input_card.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        
        tk.Label(input_card, text="📝 作文正文", font=("Microsoft YaHei", 11, "bold"), fg="#185FA5", bg="white").pack(anchor="w")
        
        # 输入方式Tab
        self.input_notebook = ttk.Notebook(input_card)
        self.input_notebook.pack(fill="both", expand=True, pady=(6, 0))
        
        # Tab1: 文字输入
        self.text_input_frame = tk.Frame(self.input_notebook, bg="white")
        self.input_notebook.add(self.text_input_frame, text="  ⌨️ 文字输入  ")
        
        text_toolbar = tk.Frame(self.text_input_frame, bg="#f0f4f8", padx=8, pady=3)
        text_toolbar.pack(fill="x")
        tk.Button(text_toolbar, text="📋 粘贴剪贴板", font=("Microsoft YaHei", 9), fg="#185FA5", bg="white", relief="solid", bd=1, padx=8, cursor="hand2", command=self._paste_from_clipboard).pack(side="left")
        tk.Label(text_toolbar, text="支持语音转文字结果直接粘贴", font=("Microsoft YaHei", 8), fg="#888888", bg="#f0f4f8").pack(side="right")
        
        self.essay_text = scrolledtext.ScrolledText(self.text_input_frame, font=("Microsoft YaHei", 11), wrap="word", bd=1, relief="solid", padx=10, pady=6, bg="#fafafa")
        self.essay_text.pack(fill="both", expand=True)
        
        # Tab2: 图片识别
        self.photo_input_frame = tk.Frame(self.input_notebook, bg="white")
        self.input_notebook.add(self.photo_input_frame, text="  📷 图片识别  ")
        
        photo_toolbar = tk.Frame(self.photo_input_frame, bg="#f0f4f8", padx=8, pady=4)
        photo_toolbar.pack(fill="x")
        self.ocr_btn = tk.Button(photo_toolbar, text="📷 上传作文照片", font=("Microsoft YaHei", 9), fg="#7B3F00", bg="#F9E79F", relief="solid", bd=1, padx=10, cursor="hand2", command=self.on_upload_photo)
        self.ocr_btn.pack(side="left")
        self.ocr_status = tk.Label(photo_toolbar, text="", font=("Microsoft YaHei", 9), fg="#888888", bg="#f0f4f8")
        self.ocr_status.pack(side="left", padx=(10, 0))
        
        self.photo_text = scrolledtext.ScrolledText(self.photo_input_frame, font=("Microsoft YaHei", 11), wrap="word", bd=1, relief="solid", padx=10, pady=8, state="disabled", bg="#fafafa")
        self.photo_text.pack(fill="both", expand=True)
        
        # Tab3: Word文档
        self.word_input_frame = tk.Frame(self.input_notebook, bg="white")
        self.input_notebook.add(self.word_input_frame, text="  📄 Word文档  ")
        
        word_toolbar = tk.Frame(self.word_input_frame, bg="#f0f4f8", padx=8, pady=4)
        word_toolbar.pack(fill="x")
        self.word_btn = tk.Button(word_toolbar, text="📂 打开Word文档", font=("Microsoft YaHei", 9), fg="#185FA5", bg="white", relief="solid", bd=1, padx=10, cursor="hand2", command=self.on_open_word)
        self.word_btn.pack(side="left")
        self.word_status = tk.Label(word_toolbar, text="", font=("Microsoft YaHei", 9), fg="#888888", bg="#f0f4f8")
        self.word_status.pack(side="left", padx=(10, 0))
        
        # Word拖拽区域
        drop_frame = tk.Frame(self.word_input_frame, bg="#E8F4FD")
        drop_frame.pack(fill="both", expand=True, padx=10, pady=(8, 5))
        
        # 创建Canvas用于绘制虚线边框（放在底层）
        border_canvas = tk.Canvas(drop_frame, bg="#E8F4FD", highlightthickness=0)
        border_canvas.pack(fill="both", expand=True)
        
        # 创建内容Frame（支持DND）
        drop_content = tk.Frame(border_canvas, bg="#E8F4FD")
        drop_content_window = border_canvas.create_window((0, 0), window=drop_content, anchor="nw")
        
        # 添加标签
        tk.Label(drop_content, text="📥 拖拽Word文档到此处", font=("Microsoft YaHei", 12), fg="#185FA5", bg="#E8F4FD").pack(pady=(30, 5))
        tk.Label(drop_content, text="或点击上方按钮选择文件", font=("Microsoft YaHei", 9), fg="#666666", bg="#E8F4FD").pack()
        if TKINTER_DND_AVAILABLE:
            tk.Label(drop_content, text="支持 .docx 格式", font=("Microsoft YaHei", 8), fg="#999999", bg="#E8F4FD").pack(pady=(10, 30))
        else:
            tk.Label(drop_content, text="（安装 tkinterdnd2 启用拖放）", font=("Microsoft YaHei", 8), fg="#CC6600", bg="#E8F4FD").pack(pady=(10, 30))
        
        # 绑定拖拽事件到内容Frame（仅在tkinterdnd2可用时）
        if TKINTER_DND_AVAILABLE:
            drop_content.drop_target_register('DND_Files')
            drop_content.dnd_bind('<<Drop>>', lambda e: self._on_word_drop(e, drop_frame))
        
        # 绘制虚线边框的函数
        def draw_dashed_border(event=None):
            border_canvas.delete("border")
            x1, y1 = 5, 5
            x2, y2 = border_canvas.winfo_width() - 5, border_canvas.winfo_height() - 5
            if x2 > x1 and y2 > y1:
                border_canvas.create_rectangle(x1, y1, x2, y2, dash=(6, 4), outline="#3498DB", width=2, tags="border")
        
        # 绑定大小变化事件
        def on_configure(event):
            border_canvas.itemconfigure(drop_content_window, width=event.width, height=event.height)
            draw_dashed_border()
        
        border_canvas.bind("<Configure>", on_configure)
        
        # 初始绘制边框
        border_canvas.after(100, draw_dashed_border)
        
        self.word_text = scrolledtext.ScrolledText(self.word_input_frame, font=("Microsoft YaHei", 11), wrap="word", bd=1, relief="solid", padx=10, pady=8, state="disabled", bg="#fafafa")
        self.word_text.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        
        # ④ 批量处理区
        batch_card = tk.Frame(left_frame, bg="white", padx=12, pady=10)
        batch_card.pack(fill="x", padx=5, pady=(0, 5))
        
        tk.Label(batch_card, text="📁 批量批改", font=("Microsoft YaHei", 11, "bold"), fg="#185FA5", bg="white").pack(anchor="w")
        
        batch_btn_row = tk.Frame(batch_card, bg="white")
        batch_btn_row.pack(fill="x", pady=(8, 5))
        
        self.batch_upload_btn = tk.Button(batch_btn_row, text="📂 上传照片", font=("Microsoft YaHei", 9), fg="white", bg="#3498DB", relief="flat", padx=10, pady=6, cursor="hand2", command=self.on_batch_upload)
        self.batch_upload_btn.pack(side="left")
        self.batch_ocr_btn = tk.Button(batch_btn_row, text="🔎 OCR识别", font=("Microsoft YaHei", 9), fg="#7B3F00", bg="#F9E79F", relief="solid", bd=1, padx=8, pady=6, cursor="hand2", command=self.on_batch_ocr, state="disabled")
        self.batch_ocr_btn.pack(side="left", padx=(5, 0))
        self.batch_grade_btn = tk.Button(batch_btn_row, text="🚀 批量批改", font=("Microsoft YaHei", 9, "bold"), fg="white", bg="#E74C3C", relief="flat", padx=8, pady=6, cursor="hand2", command=self.on_batch_grade, state="disabled")
        self.batch_grade_btn.pack(side="left", padx=(5, 0))
        self.batch_export_btn = tk.Button(batch_btn_row, text="💾 导出", font=("Microsoft YaHei", 9), fg="#185FA5", bg="white", relief="solid", bd=1, padx=8, pady=6, cursor="hand2", command=self.on_batch_export, state="disabled")
        self.batch_export_btn.pack(side="left", padx=(5, 0))
        
        list_frame = tk.Frame(batch_card, bg="#EEEEEE", height=60)
        list_frame.pack(fill="x")
        list_frame.pack_propagate(False)
        self.batch_listbox = tk.Listbox(list_frame, font=("Microsoft YaHei", 9), bg="white", bd=1, relief="solid", selectmode="extended")
        self.batch_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(list_frame, command=self.batch_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.batch_listbox.config(yscrollcommand=scrollbar.set)
        self.batch_status = tk.Label(batch_card, text="", font=("Microsoft YaHei", 9), fg="#888888", bg="white")
        self.batch_status.pack(anchor="w", pady=(3, 0))
        
        # ===== 右侧：结果区 =====
        right_frame = tk.Frame(main_paned, bg="white", padx=15, pady=12)
        main_paned.add(right_frame, weight=2)  # 右侧占40%
        
        # 结果标题栏
        result_header = tk.Frame(right_frame, bg="white")
        result_header.pack(fill="x", pady=(0, 8))
        tk.Label(result_header, text="📊 批改结果", font=("Microsoft YaHei", 13, "bold"), fg="#185FA5", bg="white").pack(side="left")
        
        # 快捷操作按钮
        quick_actions = tk.Frame(result_header, bg="white")
        quick_actions.pack(side="right")
        tk.Button(quick_actions, text="📈 班级报告", font=("Microsoft YaHei", 9), fg="white", bg="#9B59B6", relief="flat", padx=10, pady=3, cursor="hand2", command=self._generate_class_report).pack(side="left", padx=(5, 0))
        tk.Button(quick_actions, text="📂 学生报告", font=("Microsoft YaHei", 9), fg="white", bg="#3498DB", relief="flat", padx=10, pady=3, cursor="hand2", command=self._open_student_reports).pack(side="left", padx=(5, 0))
        tk.Button(quick_actions, text="📊 所有成绩", font=("Microsoft YaHei", 9), fg="white", bg="#27AE60", relief="flat", padx=10, pady=3, cursor="hand2", command=self._show_all_scores_window).pack(side="left", padx=(5, 0))
        tk.Button(quick_actions, text="📋 历史记录", font=("Microsoft YaHei", 9), fg="white", bg="#185FA5", relief="flat", padx=10, pady=3, cursor="hand2", command=self._show_history_window).pack(side="left", padx=(5, 0))
        
        self.result_text = scrolledtext.ScrolledText(right_frame, font=("Microsoft YaHei", 10), wrap="word", state="disabled", bd=0, bg="#FAFAF8", padx=12, pady=10, highlightbackground="#E0E0E0", highlightthickness=1)
        self.result_text.pack(fill="both", expand=True)

        self._set_result_text("💡 欢迎使用小学语文作文智能批改助手 v5.3\n\n【使用指南】\n\n🎯 第一步：在上方输入学生姓名和作文题目\n\n📝 文字输入：\n① 切换到「文字输入」标签\n② 直接粘贴或输入作文内容\n③ 点击绿色的「🚀 开始批改」按钮\n\n📷 图片识别：\n① 切换到「图片识别」标签\n② 点击「上传作文照片」\n③ AI自动识别文字后，点击「🚀 开始批改」\n\n📄 Word文档：\n① 切换到「Word文档」标签\n② 拖拽或点击选择 .docx 文件\n③ 内容自动导入，点击「🚀 开始批改」\n\n📁 批量批改：在「批量批改」区上传多张照片即可批量处理\n\n📊 成绩查询：点击右上角「所有成绩」查看所有学生评分\n\n💡 支持微信/讯飞/搜狗语音转文字结果直接粘贴！")

    def _show_all_scores_window(self):
        """显示所有成绩查询窗口"""
        win = tk.Toplevel(self.root)
        win.title("📊 所有学生成绩查询")
        win.geometry("1200x750")
        win.transient(self.root)
        win.grab_set()
        
        # 标题栏
        header = tk.Frame(win, bg="#185FA5", padx=15, pady=12)
        header.pack(fill="x")
        tk.Label(header, text="📊 学生作文成绩总览", font=("Microsoft YaHei", 16, "bold"), fg="white", bg="#185FA5").pack(side="left")
        
        # 加载数据
        history_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
        records = []
        if os.path.exists(history_dir):
            for f in glob.glob(os.path.join(history_dir, "*.json")):
                try:
                    with open(f, 'r', encoding='utf-8') as fp:
                        data = json.load(fp)
                        # 添加修改时间
                        data['_file'] = os.path.basename(f)
                        data['_mtime'] = os.path.getmtime(f)
                        records.append(data)
                except Exception as e:
                    print(f"读取文件失败 {f}: {e}")
                    pass
        
        if not records:
            tk.Label(win, text="📭 暂无批改记录\n请先进行作文批改", font=("Microsoft YaHei", 14), fg="#888888", bg="white").pack(expand=True)
            return
        
        # 计算统计数据
        total_students = len(records)
        total_score = 0
        score_count = 0
        
        for r in records:
            try:
                score = float(r.get('total_score', 0))
                total_score += score
                score_count += 1
            except:
                pass
        
        avg_score = total_score / score_count if score_count > 0 else 0
        
        # 统计信息栏
        stats_frame = tk.Frame(win, bg="#F0F8FF", padx=15, pady=10)
        stats_frame.pack(fill="x")
        
        stats_left = tk.Frame(stats_frame, bg="#F0F8FF")
        stats_left.pack(side="left")
        tk.Label(stats_left, text=f"📝 总记录数: {total_students}", font=("Microsoft YaHei", 11), fg="#185FA5", bg="#F0F8FF").pack(side="left", padx=(0, 30))
        tk.Label(stats_left, text=f"👥 已评分: {score_count}", font=("Microsoft YaHei", 11), fg="#27AE60", bg="#F0F8FF").pack(side="left", padx=(0, 30))
        tk.Label(stats_left, text=f"📈 平均分: {avg_score:.1f}", font=("Microsoft YaHei", 11), fg="#E74C3C", bg="#F0F8FF").pack(side="left")
        
        # 搜索框
        search_frame = tk.Frame(win, bg="white", padx=15, pady=8)
        search_frame.pack(fill="x")
        tk.Label(search_frame, text="🔍", font=("Microsoft YaHei", 12), fg="#666666", bg="white").pack(side="left")
        search_entry = tk.Entry(search_frame, font=("Microsoft YaHei", 11), bd=1, relief="solid")
        search_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # 创建Treeview显示成绩列表
        tree_frame = tk.Frame(win, bg="white", padx=15, pady=10)
        tree_frame.pack(fill="both", expand=True)
        
        # 添加滚动条
        scrollbar_y = tk.Scrollbar(tree_frame, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x = tk.Scrollbar(tree_frame, orient="horizontal")
        scrollbar_x.pack(side="bottom", fill="x")
        
        columns = ("姓名", "作文题目", "总分", "内容", "流畅", "结构", "字词", "批改时间")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", 
                           yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set,
                           selectmode="browse")
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # 设置列标题和宽度
        col_widths = {"姓名": 90, "作文题目": 180, "总分": 60, "内容": 60, 
                     "流畅": 60, "结构": 60, "字词": 60, "批改时间": 140}
        for col in columns:
            tree.heading(col, text=col)
            anchor_val = "center" if col != "作文题目" else "w"
            tree.column(col, width=col_widths[col], anchor=anchor_val)
        
        tree.pack(fill="both", expand=True)
        
        # 填充数据的函数
        def populate_tree(data_list):
            # 清空现有数据
            for item in tree.get_children():
                tree.delete(item)
            
            for r in data_list:
                # 获取各维度分数
                dims = r.get('dimensions', {})
                
                # 格式化时间
                mtime = r.get('_mtime', 0)
                if mtime:
                    from datetime import datetime
                    time_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
                else:
                    time_str = "未知"
                
                # 插入数据行
                tree.insert("", tk.END, values=(
                    r.get('student_name', '未知'),
                    r.get('essay_title', '未知'),
                    r.get('total_score', '?'),
                    dims.get('content', {}).get('score', '?'),
                    dims.get('fluency', {}).get('score', '?'),
                    dims.get('structure', {}).get('score', '?'),
                    dims.get('wrong_characters', {}).get('score', '?'),
                    time_str
                ))
        
        # 按总分排序
        def get_score(r):
            try:
                return float(r.get('total_score', 0))
            except:
                return 0
        
        sorted_records = sorted(records, key=get_score, reverse=True)
        populate_tree(sorted_records)
        
        # 搜索功能
        def on_search(*args):
            keyword = search_entry.get().strip().lower()
            if not keyword:
                populate_tree(sorted_records)
            else:
                filtered = [r for r in records if keyword in r.get('student_name', '').lower() or 
                           keyword in r.get('essay_title', '').lower()]
                populate_tree(filtered)
        
        search_entry.bind("<KeyRelease>", on_search)
        
        # 底部状态栏
        status_bar = tk.Frame(win, bg="#f0f4f8", height=40)
        status_bar.pack(fill="x", side="bottom")
        tk.Label(status_bar, text=f"共 {len(records)} 条记录  |  双击可查看详情", font=("Microsoft YaHei", 9), fg="#888888", bg="#f0f4f8").pack(side="left", padx=15)
        tk.Button(status_bar, text="🗑 清空所有成绩", font=("Microsoft YaHei", 9), 
                 fg="white", bg="#E74C3C", relief="flat", padx=12, pady=5, 
                 cursor="hand2", command=lambda: self._clear_all_scores(win)).pack(side="right", padx=15)
        
        # 双击查看详情
        def on_double_click(event):
            selection = tree.selection()
            if selection:
                item = tree.item(selection[0])
                values = item['values']
                if values:
                    student_name = values[0]
                    # 找到对应的记录
                    for r in records:
                        if r.get('student_name') == student_name:
                            self._show_grade_detail(r)
                            break
        
        tree.bind("<Double-1>", on_double_click)
    
    def _show_grade_detail(self, record):
        """显示单条批改记录的详情"""
        win = tk.Toplevel(self.root)
        name = record.get('student_name', '未知')
        title = record.get('essay_title', '未知')
        win.title(f"📊 {name} - {title} 批改详情")
        win.geometry("700x800")
        win.transient(self.root)
        win.grab_set()
        
        # 主滚动区域
        main_canvas = tk.Canvas(win, bg="white")
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(win, command=main_canvas.yview)
        scrollbar.pack(side="right", fill="y")
        main_canvas.config(yscrollcommand=scrollbar.set)
        
        content = tk.Frame(main_canvas, bg="white", padx=20, pady=20)
        main_canvas.create_window((0, 0), window=content, anchor="nw")
        
        # 标题
        tk.Label(content, text=f"📝 {name} 的作文批改详情", font=("Microsoft YaHei", 14, "bold"), fg="#185FA5", bg="white").pack(anchor="w", pady=(0, 5))
        tk.Label(content, text=f"作文题目：{title}", font=("Microsoft YaHei", 11), fg="#333333", bg="white").pack(anchor="w", pady=(0, 15))
        
        # 总分
        score_frame = tk.Frame(content, bg="#E8F6EF", padx=15, pady=12)
        score_frame.pack(fill="x", pady=(0, 15))
        tk.Label(score_frame, text="总分", font=("Microsoft YaHei", 12, "bold"), fg="#27AE60", bg="#E8F6EF").pack(side="left")
        total_score = record.get('total_score', 0)
        tk.Label(score_frame, text=f"{total_score} 分", font=("Microsoft YaHei", 24, "bold"), fg="#27AE60", bg="#E8F6EF").pack(side="left", padx=(20, 0))
        
        # 各维度详情
        dims = record.get('dimensions', {})
        dim_info = [
            ("📝 内容", 'content', dims.get('content', {})),
            ("🔤 字词", 'wrong_characters', dims.get('wrong_characters', {})),
            ("🔮 标点", 'punctuation', dims.get('punctuation', {})),
            ("💬 流畅", 'fluency', dims.get('fluency', {})),
            ("📐 结构", 'structure', dims.get('structure', {}))
        ]
        
        for dim_name, dim_key, dim_data in dim_info:
            frame = tk.Frame(content, bg="white", padx=10, pady=8)
            frame.pack(fill="x", pady=(0, 8))
            
            score = dim_data.get('score', '?')
            tk.Label(frame, text=f"{dim_name}：{score} 分", font=("Microsoft YaHei", 11, "bold"), fg="#185FA5", bg="white").pack(anchor="w")
            
            comment = dim_data.get('comment', '')
            if comment:
                tk.Label(frame, text=comment, font=("Microsoft YaHei", 9), fg="#555555", bg="white", wraplength=650, justify="left", anchor="w").pack(anchor="w", pady=(3, 0))
        
        # 亮点
        highlights = record.get('highlight_good', [])
        if highlights:
            hl_frame = tk.Frame(content, bg="#FFF9E6", padx=10, pady=8)
            hl_frame.pack(fill="x", pady=(10, 8))
            tk.Label(hl_frame, text="✨ 亮点", font=("Microsoft YaHei", 10, "bold"), fg="#F39C12", bg="#FFF9E6").pack(anchor="w")
            for hl in highlights:
                tk.Label(hl_frame, text=f"• {hl}", font=("Microsoft YaHei", 9), fg="#8B6914", bg="#FFF9E6", wraplength=650, justify="left", anchor="w").pack(anchor="w", pady=(2, 0))
        
        # 改进建议
        tip = record.get('improvement_tip', '')
        if tip:
            tip_frame = tk.Frame(content, bg="#FEF9E7", padx=10, pady=8)
            tip_frame.pack(fill="x", pady=(0, 8))
            tk.Label(tip_frame, text="💡 建议", font=("Microsoft YaHei", 10, "bold"), fg="#F39C12", bg="#FEF9E7").pack(anchor="w")
            tk.Label(tip_frame, text=tip, font=("Microsoft YaHei", 9), fg="#7B3F00", bg="#FEF9E7", wraplength=650, justify="left", anchor="w").pack(anchor="w", pady=(3, 0))
        
        # 整体评语
        overall = record.get('overall_comment', '')
        if overall:
            tk.Label(content, text="📝 整体评语", font=("Microsoft YaHei", 10, "bold"), fg="#185FA5", bg="white").pack(anchor="w", pady=(10, 3))
            tk.Label(content, text=overall, font=("Microsoft YaHei", 10), fg="#333333", bg="white", wraplength=650, justify="left", anchor="w").pack(anchor="w")
        
        # 更新滚动区域
        content.update_idletasks()
        main_canvas.config(scrollregion=main_canvas.bbox("all"))
    
    def _show_history_window(self):
        """显示历史记录窗口"""
        win = tk.Toplevel(self.root)
        win.title("📋 历史批改记录")
        win.geometry("1000x650")
        win.transient(self.root)
        win.grab_set()
        
        # 标题栏
        header = tk.Frame(win, bg="#185FA5", padx=15, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="📋 历史批改记录", font=("Microsoft YaHei", 14, "bold"), fg="white", bg="#185FA5").pack(side="left")
        
        # 操作按钮区
        action_frame = tk.Frame(win, bg="white", padx=15, pady=8)
        action_frame.pack(fill="x")
        
        tk.Button(action_frame, text="🗑 清空历史记录", font=("Microsoft YaHei", 10), 
                 fg="white", bg="#E74C3C", relief="flat", padx=12, pady=5, 
                 cursor="hand2", command=lambda: self._clear_all_history(win)).pack(side="right")
        
        # 加载数据
        history_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
        records = []
        if os.path.exists(history_dir):
            for f in glob.glob(os.path.join(history_dir, "*.json")):
                try:
                    with open(f, 'r', encoding='utf-8') as fp:
                        data = json.load(fp)
                        data['_file'] = f
                        data['_mtime'] = os.path.getmtime(f)
                        records.append(data)
                except:
                    pass
        
        if not records:
            tk.Label(win, text="📭 暂无历史记录\n请先进行作文批改", font=("Microsoft YaHei", 14), 
                     fg="#888888", bg="white").pack(expand=True)
            return
        
        # 统计信息
        stats_frame = tk.Frame(win, bg="#F0F8FF", padx=15, pady=8)
        stats_frame.pack(fill="x")
        tk.Label(stats_frame, text=f"📝 共 {len(records)} 条历史记录", 
                font=("Microsoft YaHei", 11), fg="#185FA5", bg="#F0F8FF").pack(side="left", padx=(0, 30))
        
        # 主内容区
        content = tk.Frame(win, bg="white")
        content.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        # 创建带滚动条的列表
        tree_frame = tk.Frame(content, bg="white")
        tree_frame.pack(fill="both", expand=True)
        
        scrollbar_y = tk.Scrollbar(tree_frame, orient="vertical")
        scrollbar_y.pack(side="right", fill="y")
        
        columns = ("学生姓名", "作文题目", "总分", "批改时间")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", 
                           yscrollcommand=scrollbar_y.set, selectmode="browse", height=20)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.config(command=tree.yview)
        
        # 设置列
        col_widths = {"学生姓名": 120, "作文题目": 250, "总分": 80, "批改时间": 160}
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], anchor="center")
        
        # 填充数据
        for r in sorted(records, key=lambda x: x.get('_mtime', 0), reverse=True):
            from datetime import datetime
            mtime = r.get('_mtime', 0)
            time_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M") if mtime else "未知"
            
            tree.insert("", tk.END, values=(
                r.get('student_name', '未知'),
                r.get('essay_title', '未知'),
                r.get('total_score', '?'),
                time_str
            ))
        
        # 底部状态栏
        status_bar = tk.Frame(win, bg="#f0f4f8", height=30)
        status_bar.pack(fill="x", side="bottom")
        tk.Label(status_bar, text=f"共 {len(records)} 条记录  |  双击可查看详情", 
                font=("Microsoft YaHei", 9), fg="#888888", bg="#f0f4f8").pack(side="left", padx=15)
        
        # 双击查看详情
        def on_double_click(event):
            selection = tree.selection()
            if selection:
                item = tree.item(selection[0])
                values = item['values']
                if values and len(values) >= 3:
                    student_name = values[0]
                    for r in records:
                        if r.get('student_name') == student_name and r.get('essay_title') == values[1]:
                            self._show_grade_detail(r)
                            break
        
        tree.bind("<Double-1>", on_double_click)
    
    def _clear_all_history(self, parent_window=None):
        """清空所有历史记录"""
        result = messagebox.askyesno("确认清空", 
                                     "确定要清空所有历史记录吗？\n\n此操作将删除所有批改结果文件，且不可恢复！", 
                                     icon="warning")
        if result:
            history_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
            if os.path.exists(history_dir):
                count = 0
                for f in glob.glob(os.path.join(history_dir, "*.json")):
                    try:
                        os.remove(f)
                        count += 1
                    except Exception as e:
                        print(f"删除失败 {f}: {e}")
                
                messagebox.showinfo("清空完成", f"已清空 {count} 条历史记录！")
                if parent_window:
                    parent_window.destroy()
                self._load_grade_history()  # 刷新主界面历史
            else:
                messagebox.showinfo("清空完成", "没有找到历史记录文件。")
    
    def _clear_all_scores(self, parent_window=None):
        """清空所有成绩（与清空历史记录相同操作）"""
        result = messagebox.askyesno("确认清空", 
                                     "确定要清空所有学生成绩吗？\n\n此操作将删除所有批改结果文件，且不可恢复！", 
                                     icon="warning")
        if result:
            history_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
            if os.path.exists(history_dir):
                count = 0
                for f in glob.glob(os.path.join(history_dir, "*.json")):
                    try:
                        os.remove(f)
                        count += 1
                    except Exception as e:
                        print(f"删除失败 {f}: {e}")
                
                messagebox.showinfo("清空完成", f"已清空 {count} 条成绩记录！")
                if parent_window:
                    parent_window.destroy()
                self._load_grade_history()  # 刷新主界面历史
            else:
                messagebox.showinfo("清空完成", "没有找到成绩记录文件。")
    
    def _generate_class_report(self):
        """生成班级作文分析报告"""
        try:
            result_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
            report_path, msg = generate_class_report(result_dir)
            
            if report_path:
                messagebox.showinfo("报告生成成功", 
                                  f"✅ 班级报告已生成！\n\n位置：{report_path}\n\n{messagebox.OK}")
                # 自动打开报告
                import subprocess
                subprocess.Popen(['start', '', report_path], shell=True)
            else:
                messagebox.showwarning("生成失败", f"⚠️ {msg}\n\n请先进行作文批改！")
        except Exception as e:
            messagebox.showerror("生成失败", f"生成班级报告时出错：\n{str(e)}")
    
    def _open_student_reports(self):
        """打开学生报告文件夹"""
        try:
            report_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "学生报告")
            os.makedirs(report_dir, exist_ok=True)
            
            # 检查是否有报告
            reports = glob.glob(os.path.join(report_dir, "*.html"))
            if reports:
                messagebox.showinfo("学生报告", f"📂 学生报告文件夹中共有 {len(reports)} 份报告\n\n位置：{report_dir}")
                # 自动打开文件夹
                import subprocess
                subprocess.Popen(['explorer', report_dir])
            else:
                messagebox.showinfo("学生报告", f"📭 暂无学生报告\n\n批改完成后会自动生成学生专属报告\n\n位置：{report_dir}")
        except Exception as e:
            messagebox.showerror("打开失败", f"打开文件夹时出错：\n{str(e)}")

    def _set_result_text(self, text):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", text)
        self.result_text.config(state="disabled")

    def _paste_from_clipboard(self):
        try:
            text = self.root.clipboard_get()
            if text.strip():
                self.essay_text.delete("1.0", tk.END)
                self.essay_text.insert("1.0", text.strip())
                messagebox.showinfo("粘贴成功", f"已粘贴 {len(text.strip())} 字")
            else:
                messagebox.showwarning("粘贴失败", "剪贴板为空")
        except Exception as e:
            messagebox.showwarning("粘贴失败", f"无法读取剪贴板：{str(e)}")

    def _get_current_essay_text(self):
        """根据当前Tab获取作文文本"""
        try:
            if hasattr(self, 'input_notebook') and self.input_notebook.winfo_exists():
                current_tab = self.input_notebook.index("current")
                if current_tab == 0:  # 文字输入
                    return self.essay_text.get("1.0", tk.END).strip()
                elif current_tab == 1:  # 图片识别
                    return self.photo_text.get("1.0", tk.END).strip()
                elif current_tab == 2:  # Word文档
                    return self.word_text.get("1.0", tk.END).strip()
        except:
            pass
        return self.essay_text.get("1.0", tk.END).strip()

    def on_open_word(self):
        """打开Word文档"""
        path = filedialog.askopenfilename(title="选择Word文档", filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")])
        if not path:
            return
        self._load_word_file(path)

    def _load_word_file(self, path):
        """加载Word文件内容"""
        try:
            text = read_word_text(path)
            if not text.strip():
                messagebox.showwarning("读取失败", "文档内容为空")
                return
            self.word_text.config(state="normal")
            self.word_text.delete("1.0", tk.END)
            self.word_text.insert("1.0", text)
            self.word_text.config(state="disabled")
            filename = os.path.basename(path)
            self.word_status.config(text=f"✅ 已加载：{filename}（{len(text)}字）", fg="#27AE60")
            self.word_btn.config(text="📂 重新选择")
        except ImportError:
            messagebox.showerror("缺少依赖", "请运行：pip install python-docx")
        except Exception as e:
            messagebox.showerror("读取失败", str(e))

    def _on_word_drop(self, widget, drop_frame):
        """处理Word文档拖拽"""
        try:
            files = self.root.tk.splitlist(widget.tk.call('tk_getDropFileList', self.root, widget._w))
            if files:
                file_path = files[0]
                if file_path.lower().endswith('.docx'):
                    self._load_word_file(file_path)
                else:
                    messagebox.showwarning("格式不支持", "请拖拽 .docx 格式的Word文档")
        except Exception as e:
            pass

    def _load_grade_history(self):
        history_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "批改结果")
        if not os.path.exists(history_dir):
            os.makedirs(history_dir, exist_ok=True)
            self.history_data = []
            self.filtered_history = []
            self._update_history_listbox()
            return
        json_files = glob.glob(os.path.join(history_dir, "*.json"))
        records = []
        for f in json_files:
            try:
                with open(f, 'r', encoding='utf-8') as fp:
                    data = json.load(fp)
                    data['_file'] = os.path.basename(f)
                    data['_mtime'] = os.path.getmtime(f)
                    records.append(data)
            except:
                pass
        records.sort(key=lambda x: x.get('_mtime', 0), reverse=True)
        self.history_data = records
        self.filtered_history = records
        self._update_history_listbox()

    def _update_history_listbox(self):
        self.history_listbox.delete(0, tk.END)
        for record in self.filtered_history:
            student = record.get('student_name', '未知')
            title = record.get('essay_title', '未知题目')
            score = record.get('total_score', '?')
            self.history_listbox.insert(tk.END, f"{student} · {title} · {score}分")
        self.history_count_label.config(text=f"共 {len(self.filtered_history)} 条记录")
        self._clear_detail_panel()

    def _clear_detail_panel(self):
        self.detail_title_label.config(text="请从左侧选择学生查看详情")
        self.total_score_label.config(text="总分：--", fg="#888888")
        for var in self.dimension_labels.values():
            var.set("--")
        self.highlight_text.config(text="暂无")
        self.tip_text.config(text="暂无")
        self.comment_text.config(state="normal")
        self.comment_text.delete("1.0", tk.END)
        self.comment_text.insert("1.0", "请从左侧选择学生查看详细的批改评语...")
        self.comment_text.config(state="disabled")

    def _filter_history_list(self, event=None):
        keyword = self.history_search_entry.get().strip().lower()
        if not keyword:
            self.filtered_history = self.history_data
        else:
            self.filtered_history = [r for r in self.history_data if keyword in r.get('student_name', '').lower() or keyword in r.get('essay_title', '').lower()]
        self._update_history_listbox()

    def _on_history_selected(self, event=None):
        selection = self.history_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        if idx >= len(self.filtered_history):
            return
        self._display_history_detail(self.filtered_history[idx])

    def _display_history_detail(self, record):
        student = record.get('student_name', '未知')
        title = record.get('essay_title', '未知题目')
        score = record.get('total_score', 0)
        self.detail_title_label.config(text=f"{student} · {title}")
        if score >= 90:
            color = "#27AE60"
        elif score >= 80:
            color = "#F39C12"
        elif score >= 70:
            color = "#E67E22"
        else:
            color = "#E74C3C"
        self.total_score_label.config(text=f"总分：{score}", fg=color)
        dims = record.get('dimensions', {})
        for name, score in [("错别字", dims.get('wrong_characters', {}).get('score')), ("标点符号", dims.get('punctuation', {}).get('score')), ("句子通顺", dims.get('fluency', {}).get('score')), ("段落结构", dims.get('structure', {}).get('score')), ("内容相关", dims.get('content', {}).get('score'))]:
            self.dimension_labels[name].set(str(score) if score is not None else "--")
        highlights = record.get('highlight_good', [])
        if highlights:
            self.highlight_text.config(text="\n".join(f"• {h}" for h in highlights[:3]))
        else:
            self.highlight_text.config(text="暂无")
        tip = record.get('improvement_tip', '')
        self.tip_text.config(text=tip if tip else "暂无")
        comment = record.get('overall_comment', '暂无评语')
        self.comment_text.config(state="normal")
        self.comment_text.delete("1.0", tk.END)
        self.comment_text.insert("1.0", comment)
        self.comment_text.config(state="disabled")

    def on_grade(self):
        student_name = self.name_entry.get().strip()
        essay_title = self.title_entry.get().strip()
        essay_text = self._get_current_essay_text()
        if not student_name:
            messagebox.showwarning("提示", "请输入学生姓名")
            return
        if not essay_title:
            messagebox.showwarning("提示", "请输入作文题目")
            return
        if not essay_text:
            messagebox.showwarning("提示", "请输入作文正文")
            return
        self.grade_btn.config(state="disabled", text="⏳ 批改中...")
        self.status_label.config(text="⏳ 正在批改...", fg="#E67E22")
        thread = threading.Thread(target=self._do_grade, args=(essay_text, student_name, essay_title), daemon=True)
        thread.start()

    def _do_grade(self, essay_text, student_name, essay_title):
        try:
            start_time = time.time()
            result, tokens = grade_essay_api(essay_text, student_name, essay_title)
            elapsed = time.time() - start_time
            dims = result.get("dimensions", {})
            total = result.get("total_score", "N/A")
            lines = []
            lines.append("=" * 38)
            lines.append(f"  ✅ 批改完成！耗时 {elapsed:.1f} 秒")
            lines.append("=" * 38)
            lines.append(f"  学生：{result.get('student_name','')}")
            lines.append(f"  题目：{result.get('essay_title','')}")
            lines.append(f"  🔢 总分：{total} / 100")
            lines.append("")
            lines.append("  📊 分项得分")
            lines.append("  " + "-" * 32)
            for key, label in [("wrong_characters", "错别字"), ("punctuation", "标点符号"), ("fluency", "句子通顺"), ("structure", "段落结构"), ("content", "内容相关")]:
                d = dims.get(key, {})
                lines.append(f"  {label:<8} {d.get('score','N/A'):>4}")
            lines.append("")
            lines.append("  " + "=" * 32)
            lines.append(f"  💬 整体评语：{result.get('overall_comment','')}")
            lines.append(f"  ✨ 亮点：{'；'.join(result.get('highlight_good',[]))}")
            lines.append(f"  💡 建议：{result.get('improvement_tip','')}")
            lines.append("  " + "=" * 32)
            lines.append(f"  Token消耗：{tokens}（约 ¥{tokens/1000*0.1:.3f}）")
            lines.append("=" * 38)
            output = "\n".join(lines)

            def update():
                self._set_result_text(output)
                self.grade_btn.config(state="normal", text="🚀 开始批改")
                self.status_label.config(text=f"✅ 完成 · {tokens} tokens", fg="#3B6D11")
                save_dir = os.path.dirname(os.path.abspath(__file__))
                result_dir = os.path.join(save_dir, "批改结果")
                os.makedirs(result_dir, exist_ok=True)
                txt_file = os.path.join(result_dir, f"{student_name}_结果.txt")
                with open(txt_file, "w", encoding="utf-8") as f:
                    f.write(output)
                json_file = os.path.join(result_dir, f"{student_name}_{essay_title}_批改结果.json")
                with open(json_file, "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)
                
                # 生成学生专属HTML报告
                try:
                    report_path = generate_student_report(result)
                    print(f"✅ 学生报告已生成：{report_path}")
                except Exception as e:
                    print(f"⚠️ 学生报告生成失败：{e}")
                
                self._load_grade_history()
            self.root.after(0, update)
        except Exception as e:
            def update():
                self._set_result_text(f"❌ 批改失败：\n{str(e)}")
                self.grade_btn.config(state="normal", text="🚀 开始批改")
                self.status_label.config(text="❌ 批改失败", fg="#E74C3C")
            self.root.after(0, update)

    def load_demo(self):
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, "王小明")
        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, "记一次难忘的事")
        self.essay_text.delete("1.0", tk.END)
        self.essay_text.insert("1.0", "今天天气真好，阳光明媚。妈妈带我去了公园放风筝。我先拿出风筝，然后妈妈帮我扶着线，我一边跑一边放线，风筝越飞越高，都快飞到云彩里面去了。我高兴得跳了起来！这次放风筝真开心，我下次还要来。")

    def on_clear(self):
        self.name_entry.delete(0, tk.END)
        self.title_entry.delete(0, tk.END)
        self.essay_text.delete("1.0", tk.END)
        self.photo_text.config(state="normal")
        self.photo_text.delete("1.0", tk.END)
        self.photo_text.config(state="disabled")
        self.ocr_status.config(text="")
        self.word_text.config(state="normal")
        self.word_text.delete("1.0", tk.END)
        self.word_text.config(state="disabled")
        self.word_status.config(text="")
        self.word_btn.config(text="📂 打开Word文档")
        self._set_result_text("💡 已清空，请重新输入作文信息。")

    def on_upload_photo(self):
        path = filedialog.askopenfilename(title="选择作文照片", filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"), ("所有文件", "*.*")])
        if not path:
            return
        self.ocr_path = path
        self.ocr_btn.config(state="disabled", text="⏳ 识别中...")
        self.ocr_status.config(text="⏳ 正在识别文字...", fg="#E67E22")
        thread = threading.Thread(target=self._do_ocr, args=(path,), daemon=True)
        thread.start()

    def _do_ocr(self, path):
        try:
            start = time.time()
            text = ocr_image(path)
            elapsed = time.time() - start
            def update():
                self.photo_text.config(state="normal")
                self.photo_text.delete("1.0", tk.END)
                self.photo_text.insert("1.0", text)
                self.photo_text.config(state="disabled")
                self.ocr_status.config(text=f"✅ 识别完成！耗时 {elapsed:.1f} 秒，{len(text)} 字", fg="#27AE60")
                self.ocr_btn.config(state="normal", text="📷 上传照片识别")
                if not text.strip():
                    messagebox.showwarning("识别结果", "未能从图片中识别到文字，请尝试：\n1. 确保图片清晰\n2. 确保文字方向正确\n3. 尝试调整图片亮度")
            self.root.after(0, update)
        except Exception as e:
            def update():
                self.ocr_status.config(text=f"❌ 识别失败：{str(e)}", fg="#E74C3C")
                self.ocr_btn.config(state="normal", text="📷 上传照片识别")
            self.root.after(0, update)

    def on_batch_upload(self):
        paths = filedialog.askopenfilenames(title="批量选择作文照片（可多选）", filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"), ("所有文件", "*.*")])
        if not paths:
            return
        self.batch_files = list(paths)
        self.batch_listbox.delete(0, tk.END)
        for i, path in enumerate(self.batch_files):
            self.batch_listbox.insert(tk.END, f"  {i+1}. {os.path.basename(path)}")
        self.batch_results = [{"path": p, "filename": os.path.basename(p), "ocr_text": None, "grade_result": None, "ocr_done": False, "grade_done": False} for p in self.batch_files]
        count = len(self.batch_files)
        self.batch_status.config(text=f"✅ 已加载 {count} 张照片，可点击「批量OCR识别」", fg="#27AE60")
        self.batch_ocr_btn.config(state="normal")
        self.batch_grade_btn.config(state="disabled")
        self.batch_export_btn.config(state="disabled")
        messagebox.showinfo("批量上传成功", f"已选择 {count} 张作文照片！\n\n下一步：\n1. 点击「🔎 批量OCR识别」识别所有照片\n2. 点击「🚀 批量批改全部」开始AI批改")

    def on_batch_ocr(self):
        if not self.batch_files:
            messagebox.showwarning("提示", "请先批量上传照片")
            return
        self.batch_ocr_btn.config(state="disabled", text="⏳ 批量识别中...")
        self.batch_upload_btn.config(state="disabled")
        self.batch_grade_btn.config(state="disabled")
        thread = threading.Thread(target=self._do_batch_ocr, daemon=True)
        thread.start()

    def _do_batch_ocr(self):
        total = len(self.batch_files)
        success_count = 0
        empty_count = 0
        for i, item in enumerate(self.batch_results):
            filename = item["filename"]
            self.root.after(0, lambda idx=i, fn=filename: self.batch_status.config(text=f"⏳ 正在识别 {idx+1}/{total}：{fn} ...", fg="#E67E22"))
            try:
                text = ocr_image(item["path"])
                if text.strip():
                    self.batch_results[i]["ocr_text"] = text
                    self.batch_results[i]["ocr_done"] = True
                    success_count += 1
                else:
                    empty_count += 1
            except:
                empty_count += 1
        def finish():
            self.batch_ocr_btn.config(state="normal", text="🔎 批量OCR识别")
            self.batch_upload_btn.config(state="normal")
            if success_count > 0:
                self.batch_grade_btn.config(state="normal")
                self.batch_status.config(text=f"✅ OCR完成：{success_count}张成功，{empty_count}张失败/为空", fg="#27AE60")
            else:
                self.batch_status.config(text=f"❌ OCR完成：全部失败", fg="#E74C3C")
        self.root.after(0, finish)

    def on_batch_grade(self):
        if not self.batch_files:
            messagebox.showwarning("提示", "请先批量上传照片")
            return
        dialog = tk.Toplevel(self.root)
        dialog.title("批量批改设置")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.update_idletasks()
        dialog.geometry(f"+{self.root.winfo_x()+(1200-400)//2}+{self.root.winfo_y()+(800-150)//2}")
        tk.Label(dialog, text="请输入作文题目（所有学生共用）：", font=("Microsoft YaHei", 10)).pack(pady=(20, 5))
        title_entry = tk.Entry(dialog, font=("Microsoft YaHei", 11), width=40)
        title_entry.insert(0, "记一次难忘的事")
        title_entry.pack(pady=(0, 15))

        def start_batch():
            essay_title = title_entry.get().strip()
            if not essay_title:
                messagebox.showwarning("提示", "请输入作文题目")
                return
            dialog.destroy()
            self.batch_grade_btn.config(state="disabled", text="⏳ 批改中...")
            self.batch_ocr_btn.config(state="disabled")
            self.batch_upload_btn.config(state="disabled")
            self.batch_export_btn.config(state="disabled")
            thread = threading.Thread(target=self._do_batch_grade, args=(essay_title,), daemon=True)
            thread.start()

        tk.Button(dialog, text="开始批量批改", font=("Microsoft YaHei", 11, "bold"), fg="white", bg="#27AE60", relief="flat", padx=20, pady=8, cursor="hand2", command=start_batch).pack(pady=10)

    def _do_batch_grade(self, essay_title):
        total = len(self.batch_results)
        graded_count = 0
        api_errors = 0
        for i, item in enumerate(self.batch_results):
            filename = item["filename"]
            if not item.get("ocr_text") or item["ocr_text"].startswith("[识别失败]"):
                continue
            student_name = self._extract_name_from_filename(item["filename"])
            self.root.after(0, lambda idx=i, fn=filename: self.batch_status.config(text=f"⏳ 批改中 {idx+1}/{total}：{fn} ...", fg="#E67E22"))
            try:
                result, tokens = grade_essay_api(item["ocr_text"], student_name, essay_title)
                self.batch_results[i]["grade_result"] = result
                self.batch_results[i]["student_name"] = student_name
                self.batch_results[i]["essay_title"] = essay_title
                self.batch_results[i]["grade_tokens"] = tokens
                self.batch_results[i]["grade_done"] = True
                graded_count += 1
            except:
                api_errors += 1
        def finish():
            self.batch_grade_btn.config(state="normal", text="🚀 批量批改全部")
            self.batch_ocr_btn.config(state="normal")
            self.batch_upload_btn.config(state="normal")
            if graded_count > 0:
                self.batch_export_btn.config(state="normal")
                self.batch_status.config(text=f"✅ 批改完成：{graded_count}张成功，{api_errors}张失败", fg="#27AE60")
                self._show_batch_summary()
                self._save_batch_json_results()
                self._load_grade_history()
            else:
                self.batch_status.config(text=f"❌ 批改失败：全部失败", fg="#E74C3C")
        self.root.after(0, finish)

    def _extract_name_from_filename(self, filename):
        import re
        name = os.path.splitext(filename)[0]
        name = re.sub(r'^\d+[_\-\s]*', '', name)
        name = re.sub(r'[_\-]+', ' ', name)
        return name.strip() or "学生"

    def _save_batch_json_results(self):
        save_dir = os.path.dirname(os.path.abspath(__file__))
        result_dir = os.path.join(save_dir, "批改结果")
        os.makedirs(result_dir, exist_ok=True)
        graded = [r for r in self.batch_results if r.get("grade_done")]
        for item in graded:
            student_name = item.get("student_name", "学生")
            essay_title = item.get("essay_title", "作文")
            result = item.get("grade_result")
            if result:
                json_file = os.path.join(result_dir, f"{student_name}_{essay_title}_批改结果.json")
                with open(json_file, "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=2)

    def _show_batch_summary(self):
        graded = [r for r in self.batch_results if r.get("grade_done")]
        if not graded:
            return
        summary_lines = ["=" * 45, "  📋 批量批改汇总", "=" * 45, ""]
        for i, item in enumerate(graded):
            score = item["grade_result"].get("total_score", "N/A")
            name = item.get("student_name", "未知")
            summary_lines.append(f"  {i+1}. {name}")
            summary_lines.append(f"     分数：{score} / 100")
            summary_lines.append("")
        summary_lines.append("=" * 45)
        summary_lines.append(f"  共 {len(graded)} 篇")
        summary_lines.append("=" * 45)
        self._set_result_text("\n".join(summary_lines))

    def on_batch_export(self):
        graded = [r for r in self.batch_results if r.get("grade_done")]
        if not graded:
            messagebox.showwarning("提示", "没有可导出的批改结果")
            return
        save_dir = os.path.dirname(os.path.abspath(__file__))
        result_dir = os.path.join(save_dir, "批改结果")
        os.makedirs(result_dir, exist_ok=True)
        for item in graded:
            student_name = item.get("student_name", "学生")
            essay_title = item.get("essay_title", "作文")
            result = item.get("grade_result")
            dims = result.get("dimensions", {})
            lines = []
            lines.append("=" * 40)
            lines.append(f"  学生：{student_name}")
            lines.append(f"  题目：{essay_title}")
            lines.append(f"  总分：{result.get('total_score',0)} / 100")
            lines.append("")
            for key, label in [("wrong_characters", "错别字"), ("punctuation", "标点符号"), ("fluency", "句子通顺"), ("structure", "段落结构"), ("content", "内容相关")]:
                d = dims.get(key, {})
                lines.append(f"  {label}：{d.get('score','?')}")
            lines.append("")
            lines.append(f"  评语：{result.get('overall_comment','')}")
            lines.append("=" * 40)
            txt_file = os.path.join(result_dir, f"{student_name}_{essay_title}_结果.txt")
            with open(txt_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        messagebox.showinfo("导出成功", f"已导出 {len(graded)} 份批改结果到：\n{result_dir}")

# ====================== 主程序入口 ======================
if __name__ == "__main__":
    if TKINTER_DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("提示：未安装 tkinterdnd2，拖放功能已禁用。请运行：pip install tkinterdnd2")
    app = EssayGraderApp(root)
    root.mainloop()
