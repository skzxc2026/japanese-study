import pandas as pd
import json

# 1. 读取 Excel
xls = pd.ExcelFile('data.xlsx')
df_articles = pd.read_excel(xls, 'Articles')
df_content = pd.read_excel(xls, 'Content')
df_grammar = pd.read_excel(xls, 'Grammar')

# 2. 初始化最终数据结构
final_db = {
    "writing_material_exam": [],
    "writing_material_small": [],
    "writing_material_large": []
}

# 3. 遍历每一篇文章进行组装
for index, row in df_articles.iterrows():
    art_id = row['id']
    art_type = row['type']  # exam, small, large
    
    # 构建基础对象
    article_obj = {
        "id": art_id,
        "title": row['title'],
        "year": str(row['year']),
        "requirements": row['requirements'],
        "full_cn": row['full_cn'],
        "content_struct": [],
        "grammar": []
    }

    # --- 处理正文 (分组) ---
    # 找到属于这篇文章的所有句子
    sentences = df_content[df_content['article_id'] == art_id]
    # 按段落分组
    if not sentences.empty:
        max_para = sentences['paragraph_num'].max()
        for p in range(1, int(max_para) + 1):
            para_sentences = sentences[sentences['paragraph_num'] == p]
            para_list = []
            for _, s_row in para_sentences.iterrows():
                para_list.append({
                    "text": s_row['japanese'],
                    "furigana": s_row['furigana'],
                    "audio": s_row['audio_url']
                })
            article_obj['content_struct'].append(para_list)

    # --- 处理语法 ---
    grammars = df_grammar[df_grammar['article_id'] == art_id]
    for _, g_row in grammars.iterrows():
        article_obj['grammar'].append({
            "point": g_row['point'],
            "level": g_row['level'],
            "meaning": g_row['meaning'],
            "desc": g_row['desc']
        })

    # --- 分类放入对应的列表 ---
    if art_type == 'exam':
        final_db['writing_material_exam'].append(article_obj)
    elif art_type == 'small':
        final_db['writing_material_small'].append(article_obj)
    elif art_type == 'large':
        final_db['writing_material_large'].append(article_obj)

# 4. 写入 data.js
js_content = f"window.db = {json.dumps(final_db, ensure_ascii=False, indent=2)};"
with open('assets/data.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

print("✅ 转换完成！data.js 已更新")
