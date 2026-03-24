import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
from datetime import datetime

def get_display_width(text):
    """计算字符串的显示宽度"""
    if pd.isna(text) or text is None: return 0
    text = str(text)
    return sum(2 if ord(c) > 127 else 1.2 for c in text)

def auto_adjust_column_width(file_path):
    """自动调整 Excel 列宽并设置左对齐"""
    wb = load_workbook(file_path)
    ws = wb.active
    
    from openpyxl.styles import Alignment
    
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value is not None:
                width = get_display_width(cell.value)
                if width > max_len:
                    max_len = width
                # 设置左对齐
                cell.alignment = Alignment(horizontal='left')
        ws.column_dimensions[col_letter].width = max_len + 2
        
    wb.save(file_path)

def scrape_ccass_single(file_path, start_date_str, end_date_str, stock_id="6639"):
    """单次查询：直接用开始日期和结束日期查询一次"""
    url = f"https://ccass.gogudata.com/changes/{stock_id}?startTime={start_date_str}&endTime={end_date_str}"
    print(f"正在查询 {start_date_str} 至 {end_date_str}...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,channel="msedge")
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        page.goto(url)
        
        try:
            page.wait_for_selector('table tbody tr', timeout=5000)
            rows = page.locator('table tbody tr').all()
            
            data = []
            for row in rows:
                cells = row.locator('td').all_inner_texts()
                if len(cells) >= 6:
                    data.append({
                        "序列": cells[0].strip(),
                        "席位id": cells[1].strip(),
                        "券商名称": cells[2].strip(),
                        f"{end_date_str}持股量": cells[3].strip(),
                        f"较{start_date_str}持股变动": cells[4].strip(),
                        f"{end_date_str}持股占比%": cells[5].strip()
                    })
            
            browser.close()
            
            if data:
                df = pd.DataFrame(data)
                
                # 检查文件是否存在，如果存在则追加数据
                if os.path.exists(file_path):
                    try:
                        df_existing = pd.read_excel(file_path)
                        # 打印现有数据的列名和行数，用于调试
                        print(f"现有文件包含 {len(df_existing)} 行数据，列名: {list(df_existing.columns)}")
                        print(f"新数据包含 {len(df)} 行数据，列名: {list(df.columns)}")
                        
                        # 保留原有的固定列
                        fixed_cols = ['序列', '席位id', '券商名称']
                        # 获取新数据中的非固定列（即带日期的列）
                        new_cols = [col for col in df.columns if col not in fixed_cols]
                        
                        # === 修改点：为了允许重名列且不覆盖，先创建临时列 ===
                        temp_cols = {}
                        for col in new_cols:
                            temp_name = f"__temp_{col}__"
                            df_existing[temp_name] = ""
                            temp_cols[col] = temp_name
                        
                        # 更新数据
                        for _, row in df.iterrows():
                            # 找到对应的行
                            mask = (df_existing['席位id'] == row['席位id']) & (df_existing['券商名称'] == row['券商名称'])
                            if mask.any():
                                # 更新现有行 (写入临时列，避免覆盖旧的同名列)
                                for col in new_cols:
                                    df_existing.loc[mask, temp_cols[col]] = row[col]
                            else:
                                # 添加新行
                                new_row = {c: "" for c in df_existing.columns}
                                # 填入基础信息
                                for c in fixed_cols:
                                    if c in row: new_row[c] = row[c]
                                # 填入新查到的数据
                                for col in new_cols:
                                    new_row[temp_cols[col]] = row[col]
                                
                                # 重新生成序列
                                new_row['序列'] = len(df_existing) + 1
                                df_existing = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
                        
                        # === 修改点：数据填完后，把临时列名强行改回原来的名字（允许重名） ===
                        df_existing = df_existing.rename(columns={v: k for k, v in temp_cols.items()})
                        
                        # 填充空值
                        df_existing = df_existing.fillna("")
                        
                        print(f"更新后包含 {len(df_existing)} 行数据，列名: {list(df_existing.columns)}")
                        df = df_existing
                    except Exception as e:
                        print(f"读取现有文件失败: {e}")
                        # 如果读取失败，使用新数据
                else:
                    print("文件不存在，创建新文件")
                
                df.to_excel(file_path, index=False)
                auto_adjust_column_width(file_path)
                print(f"✅ 查询成功！共 {len(data)} 条数据，已保存至 {file_path}")
                return True
            else:
                print("无数据")
                return False
                
        except Exception as e:
            browser.close()
            print(f"查询失败: {e}")
            return False

def scrape_ccass_horizontal(file_path, start_date_str, end_date_str, stock_id="6639"):
    # 如果本地已经有表格，读取它作为主表；如果没有，创建一个空表
    if os.path.exists(file_path):
        try:
            df_main = pd.read_excel(file_path)
        except Exception as e:
            print(f"读取原有 Excel 失败: {e}")
            df_main = pd.DataFrame()
    else:
        df_main = pd.DataFrame()

    # 生成日期序列
    date_list = pd.date_range(start=start_date_str, end=end_date_str, freq='D')

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print(f"开始横向排版抓取，总跨度：{start_date_str} 到 {end_date_str}")
        
        for i in range(len(date_list) - 1):
            s_date = date_list[i].strftime("%Y-%m-%d")
            e_date = date_list[i+1].strftime("%Y-%m-%d")
            
            # 为表头生成带日期的前缀，例如 "2025年3月2日"
            e_date_obj = date_list[i+1]
            date_prefix = f"{e_date_obj.year}年{e_date_obj.month}月{e_date_obj.day}日"

            url = f"https://ccass.gogudata.com/changes/{stock_id}?startTime={s_date}&endTime={e_date}"
            print(f"正在抓取 {e_date} ...", end=" ")
            page.goto(url)
            
            try:
                page.wait_for_selector('table tbody tr', timeout=5000)
                rows = page.locator('table tbody tr').all()
                
                daily_data = []
                for row in rows:
                    cells = row.locator('td').all_inner_texts()
                    if len(cells) >= 6:
                        daily_data.append({
                            "序列": cells[0].strip(),
                            "席位id": cells[1].strip(),
                            "券商名称": cells[2].strip(),
                            f"{date_prefix}持股量": cells[3].strip(),
                            f"{date_prefix}持股变动": cells[4].strip(),
                            f"{date_prefix}持股占比%": cells[5].strip()
                        })
                
                if daily_data:
                    df_daily = pd.DataFrame(daily_data)
                    
                    if df_main.empty:
                        # 第一天的数据，直接作为主表
                        df_main = df_daily
                    else:
                        # 后续天数，去掉"序列"列，仅保留关键合并字段和新增的3列数据
                        df_daily = df_daily.drop(columns=["序列"])
                        # 使用外连接 (outer merge) 横向拼接，确保新出现的券商也会被记录
                        df_main = pd.merge(df_main, df_daily, on=["席位id", "券商名称"], how="outer")
                    
                    print(f"成功 ({len(daily_data)} 条)，目前共 {len(df_main.columns)} 列")
                else:
                    print("无有效数据")
                    
            except Exception:
                print("跳过 (周末/节假日无数据)")
            
            time.sleep(1.5)

        browser.close()

    # 保存并整理数据
    if not df_main.empty:
        # 重置并整理最左侧的“序列”
        df_main['序列'] = range(1, len(df_main) + 1)
        
        # 确保 序列, 席位id, 券商名称 永远排在最左侧前三列
        fixed_cols = ['序列', '席位id', '券商名称']
        other_cols = [c for c in df_main.columns if c not in fixed_cols]
        df_main = df_main[fixed_cols + other_cols]

        # 如果某些券商在某些日期没有数据，Excel 里会显示为空 (NaN)，我们将其填充为空字符串
        df_main = df_main.fillna("")
        
        # 检查是否有实际的数据（除了序列、席位id和券商名称之外的列是否有值）
        has_data = False
        if other_cols:  # 确保有数据列
            for col in other_cols:
                if (df_main[col].str.strip() != "").any():
                    has_data = True
                    break
        else:
            # 没有数据列，说明没有获取到任何数据
            has_data = False
        
        if not has_data:
            print("\n抓取结束，没有获取到任何数据。")
            raise ValueError("缺少数据")

        # 确保目录存在
        output_dir = os.path.dirname(file_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        try:
            df_main.to_excel(file_path, index=False)
            print(f"\n✅ 数据更新完毕！已保存至: {file_path}")
        except PermissionError:
            raise PermissionError(f"文件正在被其他程序打开，请关闭后再试")
        
        print("正在自动排版并调整列宽...")
        auto_adjust_column_width(file_path)
        print("排版完成！")
    else:
        print("\n抓取结束，没有获取到任何数据。")
        raise ValueError("缺少数据")

if __name__ == "__main__":
    FILE_PATH = r"C:\Users\IT\Desktop\券商\券商数据.xlsx"
    
    # 从 2025.3.1 开始
    START_DATE = "2026-03-01"
    # 一直抓取到今天 2026.3.3
    END_DATE = "2026-03-02" 
    
    scrape_ccass_horizontal(FILE_PATH, START_DATE, END_DATE)