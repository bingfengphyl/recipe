import pandas as pd
import os
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader
import json


def excel_date_to_string(serial_date):
    try:
        if isinstance(serial_date, (int, float)):
            start = pd.to_datetime('1899-12-30')
            return (start + pd.Timedelta(days=serial_date)).strftime('%Y年%m月%d日')
        elif isinstance(serial_date, str):
            return serial_date
        else:
            return serial_date.strftime('%Y年%m月%d日')
    except Exception as e:
        raise ValueError(f"无法解析日期: {serial_date}, 错误: {e}")
    
def format_date_jinja(value):
    try:
        if isinstance(value, str) and '年' in value:
            return value.replace('年', '-').replace('月', '-').replace('日', '')
        elif isinstance(value, (int, float)):
            start = pd.to_datetime('1899-12-30')
            dt = start + pd.Timedelta(days=value)
            return dt.strftime('%Y-%m-%d')
        else:
            return value.strftime('%Y-%m-%d')
    except Exception as e:
        return '2025-05-28' 

# 设置 Jinja2 模板环境
template_dir = r'D:\02github\recipe'
env = Environment(loader=FileSystemLoader(template_dir))
env.filters['format_date'] = format_date_jinja
template_file = '猪娃家一日三餐食谱 - 示例动态.html'
output_dir = os.path.join(template_dir, 'docs')

os.makedirs(output_dir, exist_ok=True)

# 读取 Excel 数据
excel_path = r'D:\02github\recipe\\猪娃家的食谱记录.xlsx'
df = pd.read_excel(excel_path, sheet_name='Sheet1')  # 假设数据在 Sheet1 中

# 按日期分组
grouped = df.groupby('日期')
formatted_dates = []

for date, group in grouped:
    meals = {
        '早餐': [],
        '午餐': [],
        '晚餐': []
    }

    for _, row in group.iterrows():
        meal_type = row['早中晚餐别']
        link = row.get('菜谱链接', '')
        if pd.isna(link) or str(link).strip().lower() in ['nan', 'none']:
            link = ''
        dish = {
            'name': row['菜谱名称'],
            'desc': row.get('菜谱描述', ''),
            'link': link,
        }
        meals[meal_type].append(dish)


    formatted_date = excel_date_to_string(date)
    formatted_dates.append(formatted_date.replace('年', '-').replace('月', '-').replace('日', ''))

    # 使用 Jinja 渲染模板
    template = env.get_template(template_file)
    rendered_html = template.render(
        date=formatted_date,
        breakfast=meals['早餐'],
        lunch=meals['午餐'],
        dinner=meals['晚餐']
    )

    # 写入文件
    safe_date = formatted_date.replace('年', '-').replace('月', '-').replace('日', '')
    output_file = os.path.join(output_dir, f'{safe_date}.html')
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(rendered_html)

    print(f"✅ 已生成: {output_file}")

# 获取 minDate 和 maxDate
min_date = formatted_dates[0] if formatted_dates else "2025-01-01"
max_date = formatted_dates[-1] if formatted_dates else "2025-12-31"

# 构建 JSON 数据
date_data = {
    "dates": formatted_dates,
    "minDate": min_date,
    "maxDate": max_date
}
json_output_path = os.path.join(output_dir, 'dates.json')
try:
    with open(json_output_path, 'w', encoding='utf-8') as f:
        json.dump(date_data, f, ensure_ascii=False, indent=2)
    print(f"✅ JSON生成：{json_output_path}")
except Exception as e:
    print(f"❌ JSON失败，错误信息：{e}")
# if __name__ == '__main__':
#     template_path = r'D:\02github\practice-demo\20250528_HTML\猪娃家一日三餐食谱 - 示例动态.html'
#     if not os.path.exists(template_path):
#         raise FileNotFoundError(f"找不到模板文件：{template_path}")
#     else:
#         print("✅ 模板文件存在")

