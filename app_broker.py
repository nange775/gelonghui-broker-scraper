import os
import sys
import json
import pandas as pd
from flask import Flask, request, jsonify, render_template
import fetch as fetch_module
# 新增：处理打包后资源路径的工具函数
def resource_path(relative_path):
    """获取资源绝对路径，兼容开发环境和 PyInstaller 打包后的环境"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 会将资源解压到这个临时文件夹
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# 获取 EXE 所在目录（用于保存 settings.json 和 Excel）
if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

# 初始化 Flask，明确指定本地化的模板和静态文件夹
app = Flask(__name__, 
            template_folder=resource_path('templates'),
            static_folder=resource_path('static'))

SETTINGS_FILE = os.path.join(current_dir, 'settings.json')

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {'file_path': r"C:\Users\IT\Desktop\券商\券商数据.xlsx", 'stock_id': '6639'}

def save_settings(settings):
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

@app.route('/')
def index():
    return render_template('broker.html')

@app.route('/api/settings', methods=['GET'])
def get_settings():
    settings = load_settings()
    return jsonify(settings)

@app.route('/api/settings', methods=['POST'])
def update_settings():
    settings = request.json
    save_settings(settings)
    return jsonify({'success': True})

@app.route('/api/broker', methods=['POST'])
def get_broker_data():
    start_date = request.json.get('start_date')
    end_date = request.json.get('end_date')
    file_path = request.json.get('file_path')
    stock_id = request.json.get('stock_id', '6639')  # 默认股票ID为6639
    
    if not start_date or not end_date:
        return jsonify({'success': False, 'message': '开始日期和结束日期都不能为空'})
    
    # 如果没有提供文件路径，使用默认值
    if not file_path:
        file_path = r"C:\Users\IT\Desktop\券商\券商数据.xlsx"
    
    try:
        # 使用scrape_ccass_single函数，它会追加数据而不是覆盖
        success = fetch_module.scrape_ccass_single(file_path, start_date, end_date, stock_id)
        
        # 保存设置
        save_settings({'file_path': file_path, 'stock_id': stock_id})
        
        # 读取数据并返回
        if success:
            # 读取所有列作为字符串，避免pandas将空值转换为NaN
            df = pd.read_excel(file_path, dtype=str)
            # 替换任何NaN值为为空字符串
            data = df.fillna('').to_dict('records')
            return jsonify({'success': True, 'message': f'数据已保存至 {file_path}', 'data': data})
        else:
            raise ValueError("缺少数据")
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True, port=5001)
